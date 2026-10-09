import json
import subprocess
import sys

from openpyxl import Workbook


def test_regrade_cli_keeps_timeout_and_raw_record_and_checks_hashes(tmp_path) -> None:
    from zalos_sheetops.benchmarks.public.fixtures import file_sha256

    dataset = tmp_path / "data"
    dataset.mkdir()
    book = Workbook()
    book.active["A1"] = 3
    original, golden = dataset / "initial.xlsx", dataset / "golden.xlsx"
    book.save(original)
    book.save(golden)
    root = tmp_path / "agents"
    workspace = root / "runs/1/baseline/0/workspace"
    workspace.mkdir(parents=True)
    output = workspace / "output.xlsx"
    book.save(output)
    tasks = tmp_path / "tasks.json"
    tasks.write_text(
        json.dumps(
            [
                {
                    "id": 1,
                    "spreadsheet_path": ".",
                    "answer_position": "A1",
                    "instruction": "Set A1 to 3",
                }
            ]
        )
    )
    raw = root / "agents.jsonl"
    record = {
        "task": 1,
        "arm": "baseline",
        "repeat": 0,
        "seconds": 180,
        "executionStatus": "timeout",
        "grading": {"status": "timeout"},
        "usage": {},
        "inputSha256": file_sha256(original),
        "goldenSha256": file_sha256(golden),
        "outputSha256": file_sha256(output),
    }
    raw.write_text(json.dumps(record) + "\n")
    before = raw.read_bytes()
    destination = tmp_path / "graded.jsonl"
    command = [
        sys.executable,
        "-m",
        "zalos_sheetops.benchmarks.public.regrade",
        "--dataset",
        str(dataset),
        "--tasks",
        str(tasks),
        "--runs",
        str(root),
        "--output",
        str(destination),
    ]
    result = subprocess.run(command, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert json.loads(result.stdout)["recordCount"] == 1
    graded = json.loads(destination.read_text())
    assert graded["grading"]["status"] == "timeout"
    assert graded["grading"]["workbookStatus"] == "pass"
    assert graded["graderSha256"]
    assert raw.read_bytes() == before
    book.active["A1"] = 99
    book.save(original)
    result = subprocess.run(command, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert json.loads(destination.read_text())["grading"]["status"] == "source_mutated"
