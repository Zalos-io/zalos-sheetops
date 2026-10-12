import json
import subprocess
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from openpyxl import Workbook


def test_native_copy_cli_requires_calculated_caches_and_keeps_sources(tmp_path) -> None:
    from zalos_sheetops.benchmarks.public.fixtures import file_sha256

    dataset = tmp_path / "data"
    dataset.mkdir()
    original, golden = dataset / "initial.xlsx", dataset / "golden.xlsx"
    book = Workbook()
    book.active["A1"] = 1
    book.save(original)
    book.active["A1"] = "=1+2"
    book.save(golden)
    runs = tmp_path / "agents"
    workspace = runs / "runs/1/baseline/0/workspace"
    workspace.mkdir(parents=True)
    actual = workspace / "output.xlsx"
    book.save(actual)
    tasks = tmp_path / "tasks.json"
    tasks.write_text(json.dumps([{"id": 1, "spreadsheet_path": "."}]))
    record = {
        "task": 1,
        "arm": "baseline",
        "repeat": 0,
        "inputSha256": file_sha256(original),
        "goldenSha256": file_sha256(golden),
        "outputSha256": file_sha256(actual),
    }
    (runs / "agents.jsonl").write_text(json.dumps(record) + "\n")
    output = tmp_path / "native"
    command = [sys.executable, "-m", "zalos_sheetops.benchmarks.public.native"]
    result = subprocess.run(
        [
            *command,
            "prepare",
            "--dataset",
            str(dataset),
            "--tasks",
            str(tasks),
            "--runs",
            str(runs),
            "--output",
            str(output),
            "--excel-version",
            "test-fixture",
        ],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    summary = json.loads(result.stdout)
    assert summary["queueCount"] == 2
    pending = subprocess.run(
        [*command, "verify", "--output", str(output)], capture_output=True, text=True
    )
    assert json.loads(pending.stdout)["ready"] is False
    manifest = json.loads((output / "manifest.json").read_text())
    before = {str(p): file_sha256(p) for p in [original, golden, actual]}
    for entry in manifest["files"]:
        from pathlib import Path

        copy = Path(entry["copy"])
        with ZipFile(copy) as archive:
            members = {name: archive.read(name) for name in archive.namelist()}
        xml = ET.fromstring(members["xl/worksheets/sheet1.xml"])
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        xml.find(".//" + ns + "v").text = "3"
        members["xl/worksheets/sheet1.xml"] = ET.tostring(xml)
        with ZipFile(copy, "w") as archive:
            for name, data in members.items():
                archive.writestr(name, data)
    result = subprocess.run(
        [*command, "verify", "--output", str(output)], capture_output=True, text=True
    )
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["ready"] is True
    assert before == {str(p): file_sha256(p) for p in [original, golden, actual]}
