"""Prepare isolated native-Excel evaluation copies and audit saved formula caches."""

import argparse
import json
import shutil
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZipFile

from .agents import task_paths
from .fixtures import file_sha256
from .report import read_records
from .rubric import ORACLE_ISSUES


def formula_caches(path: Path) -> dict:
    formulas = missing = errors = 0
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.startswith("xl/worksheets/") or not name.endswith(".xml"):
                continue
            for cell in ET.fromstring(archive.read(name)).iter(ns + "c"):
                if cell.find(ns + "f") is None:
                    continue
                formulas += 1
                cache = cell.find(ns + "v")
                if cache is None or (not cache.text and cell.get("t") != "str"):
                    missing += 1
                errors += cell.get("t") == "e"
    return {"formulaCells": formulas, "missingCaches": missing, "errorCaches": errors}


def prepare_native_inputs(
    dataset: Path, tasks: list[dict], runs: Path, output: Path, excel_version: str
) -> dict:
    output.mkdir(parents=True, exist_ok=True)
    manifest_path = output / "manifest.json"
    if manifest_path.exists():
        manifest = json.loads(manifest_path.read_text())
        if manifest["excelVersion"] != excel_version:
            raise ValueError("Native Excel version differs from existing evaluation")
    else:
        manifest = {
            "engine": "Microsoft Excel",
            "excelVersion": excel_version,
            "createdAt": datetime.now(timezone.utc).isoformat(),
            "files": [],
        }
    known = {entry["copy"]: entry for entry in manifest["files"]}
    records = read_records(runs / "agents.jsonl")
    by_id = {str(t["id"]): t for t in tasks}
    sources = {}
    for task in tasks:
        if str(task["id"]) in ORACLE_ISSUES:
            continue
        _, golden = task_paths(dataset, task)
        copy = output / str(task["id"]) / "golden" / golden.name
        sources[str(copy)] = {
            "task": task["id"],
            "role": "golden",
            "source": str(golden),
            "sourceSha256": file_sha256(golden),
        }
    for record in records:
        if str(record["task"]) in ORACLE_ISSUES:
            continue
        task = by_id[str(record["task"])]
        initial, golden = task_paths(dataset, task)
        source = (
            runs
            / "runs"
            / str(task["id"])
            / record["arm"]
            / str(record["repeat"])
            / "workspace/output.xlsx"
        )
        for key, path in [
            ("inputSha256", initial),
            ("goldenSha256", golden),
            ("outputSha256", source),
        ]:
            actual_hash = file_sha256(path) if path.exists() else None
            if key in record and record[key] != actual_hash:
                raise ValueError(f"Recorded source mutated: {record['task']} {key}")
        if not source.exists():
            continue
        copy = (
            output
            / str(task["id"])
            / record["arm"]
            / str(record["repeat"])
            / "output.xlsx"
        )
        sources[str(copy)] = {
            "task": task["id"],
            "arm": record["arm"],
            "repeat": record["repeat"],
            "role": "actual",
            "source": str(source),
            "sourceSha256": file_sha256(source),
        }
    for filename, entry in sources.items():
        if filename in known:
            if known[filename]["sourceSha256"] != entry["sourceSha256"]:
                raise ValueError("Native evaluation source changed")
            continue
        copy = Path(filename)
        if copy.exists():
            raise ValueError("Untracked native evaluation copy already exists")
        copy.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(entry["source"], copy)
        manifest["files"].append(
            {
                **entry,
                "copy": str(copy),
                "copyInitialSha256": file_sha256(copy),
                "initialFormulaCells": formula_caches(copy)["formulaCells"],
            }
        )
    manifest_path.write_text(json.dumps(manifest, indent=2))
    queue = [
        entry
        for entry in manifest["files"]
        if entry["initialFormulaCells"]
        and file_sha256(Path(entry["copy"])) == entry["copyInitialSha256"]
    ]
    (output / "queue.json").write_text(json.dumps(queue, indent=2))
    return {
        "fileCount": len(manifest["files"]),
        "queueCount": len(queue),
        "excelVersion": excel_version,
    }


def verify_native_inputs(output: Path) -> dict:
    manifest = json.loads((output / "manifest.json").read_text())
    files = []
    for entry in manifest["files"]:
        copy = Path(entry["copy"])
        stats = formula_caches(copy)
        source_intact = file_sha256(Path(entry["source"])) == entry["sourceSha256"]
        saved = file_sha256(copy) != entry["copyInitialSha256"]
        ready = source_intact and (
            not entry["initialFormulaCells"]
            or (
                saved
                and not stats["missingCaches"]
                and stats["formulaCells"] >= entry["initialFormulaCells"]
            )
        )
        files.append(
            {
                **entry,
                **stats,
                "sourceIntact": source_intact,
                "saved": saved,
                "ready": ready,
                "evaluatedSha256": file_sha256(copy),
            }
        )
    result = {
        "engine": manifest["engine"],
        "excelVersion": manifest["excelVersion"],
        "ready": all(f["ready"] for f in files),
        "pendingCount": sum(not f["ready"] for f in files),
        "files": files,
    }
    (output / "audit.json").write_text(json.dumps(result, indent=2))
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["prepare", "verify"])
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--dataset", type=Path)
    parser.add_argument("--tasks", type=Path)
    parser.add_argument("--runs", type=Path)
    parser.add_argument("--excel-version")
    args = parser.parse_args()
    if args.action == "prepare":
        result = prepare_native_inputs(
            args.dataset,
            json.loads(args.tasks.read_text()),
            args.runs,
            args.output,
            args.excel_version,
        )
    else:
        audited = verify_native_inputs(args.output)
        result = {k: v for k, v in audited.items() if k != "files"}
    print(json.dumps(result))


if __name__ == "__main__":
    main()
