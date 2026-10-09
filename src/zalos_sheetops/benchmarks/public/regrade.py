"""Grade saved outputs independently of inference; keep original records intact."""

import argparse
import json
from collections import Counter
from pathlib import Path

from .agents import extract_usage, task_paths
from .fixtures import file_sha256
from .grading import grade_workbook
from .report import read_records


def regrade_records(
    dataset: Path,
    tasks: list[dict],
    runs: Path,
    destination: Path,
    soffice: Path | None = None,
) -> dict:
    raw_path = runs / "agents.jsonl"
    if destination.resolve() == raw_path.resolve():
        raise ValueError("Regrading must preserve the original records")
    by_id = {str(t["id"]): t for t in tasks}
    records = read_records(raw_path)
    graded = []
    grader_hash = file_sha256(Path(__file__).with_name("grading.py"))
    for record in records:
        task = by_id[str(record["task"])]
        original, golden = task_paths(dataset, task)
        root = (
            runs / "runs" / str(record["task"]) / record["arm"] / str(record["repeat"])
        )
        output = root / "workspace/output.xlsx"
        hashes = {
            "inputSha256": file_sha256(original),
            "goldenSha256": file_sha256(golden),
            "outputSha256": file_sha256(output) if output.exists() else None,
        }
        mutated = [
            key
            for key, value in hashes.items()
            if key in record and record[key] != value
        ]
        cached = (
            root / "evaluation/golden" / golden.name,
            root / "evaluation/actual/output.xlsx",
        )
        evaluated = cached if all(p.exists() for p in cached) else None
        grading = (
            {"status": "source_mutated", "hashFields": mutated}
            if mutated
            else grade_workbook(
                original,
                golden,
                output,
                task,
                soffice,
                root / "evaluation",
                evaluated=evaluated,
            )
        )
        execution_status = record.get("executionStatus", "completed")
        if not mutated and execution_status != "completed":
            grading = {
                **grading,
                "workbookStatus": grading["status"],
                "status": execution_status,
            }
        usage = extract_usage(root / "home", root / "usage.json")
        graded.append(
            {
                **record,
                **hashes,
                "originalGrading": record["grading"],
                "grading": grading,
                "graderSha256": grader_hash,
                "evaluationSha256": {
                    p.parent.name: file_sha256(p) for p in cached if p.exists()
                },
                "usage": {**record.get("usage", {}), **usage},
            }
        )
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(
        "".join(json.dumps(r, ensure_ascii=False) + "\n" for r in graded),
        encoding="utf-8",
    )
    return {
        "recordCount": len(graded),
        "statusCounts": dict(Counter(r["grading"]["status"] for r in graded)),
        "rawRecordsSha256": file_sha256(raw_path),
        "graderSha256": grader_hash,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dataset", type=Path, required=True)
    parser.add_argument("--tasks", type=Path, required=True)
    parser.add_argument("--runs", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--soffice", type=Path)
    args = parser.parse_args()
    result = regrade_records(
        args.dataset,
        json.loads(args.tasks.read_text()),
        args.runs,
        args.output,
        args.soffice,
    )
    print(json.dumps(result))


if __name__ == "__main__":
    main()
