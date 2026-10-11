import json
from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.benchmarks.public.regrade import regrade_records


def test_correct_staged_workbook_without_receipt_remains_blocked(
    tmp_path: Path,
) -> None:
    dataset = tmp_path / "data"
    dataset.mkdir()
    book = Workbook()
    book.active["A1"] = 3
    book.save(dataset / "initial.xlsx")
    book.save(dataset / "golden.xlsx")
    runs = tmp_path / "agents"
    workspace = runs / "runs/1/baseline/0/workspace"
    workspace.mkdir(parents=True)
    book.save(workspace / "output.xlsx")
    record = {
        "task": 1,
        "arm": "baseline",
        "repeat": 0,
        "executionStatus": "completed",
        "grading": {"status": "delivery_blocked"},
        "delivery": {"status": "blocked"},
        "usage": {},
    }
    raw = runs / "agents.jsonl"
    raw.write_text(json.dumps(record) + "\n")
    destination = tmp_path / "graded.jsonl"
    regrade_records(
        dataset,
        [
            {
                "id": 1,
                "spreadsheet_path": ".",
                "answer_position": "A1",
                "instruction": "Set A1 to 3",
            }
        ],
        runs,
        destination,
    )
    graded = json.loads(destination.read_text())
    assert graded["grading"]["status"] == "delivery_blocked"
    assert graded["grading"]["workbookStatus"] == "pass"
    assert json.loads(raw.read_text()) == record
