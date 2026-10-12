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
    from zalos_sheetops.benchmarks.public.native import (
        prepare_native_inputs,
        verify_native_inputs,
    )

    native = tmp_path / "native"
    prepare_native_inputs(dataset, json.loads(tasks.read_text()), root, native, "test")
    verify_native_inputs(native)
    native_command = [*command, "--native-evaluation-root", str(native)]
    result = subprocess.run(native_command, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    graded = json.loads(destination.read_text())
    assert graded["evaluationEngine"] == {"name": "Microsoft Excel", "version": "test"}
    assert graded["grading"]["workbookStatus"] == "pass"
    audit_path = native / "audit.json"
    audit = json.loads(audit_path.read_text())
    actual_entry = next(f for f in audit["files"] if f["role"] == "actual")
    actual_entry["nativeRejection"] = "Excel offered content recovery; declined"
    audit_path.write_text(json.dumps(audit))
    result = subprocess.run(native_command, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert (
        json.loads(destination.read_text())["grading"]["workbookStatus"]
        == "invalid_output"
    )
    (native / "1/baseline/0/output.xlsx").write_bytes(b"changed after audit")
    result = subprocess.run(native_command, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert (
        json.loads(destination.read_text())["grading"]["status"]
        == "evaluation_unavailable"
    )
    book.active["A1"] = 99
    book.save(original)
    result = subprocess.run(command, capture_output=True, text=True)
    assert result.returncode == 0, result.stderr
    assert json.loads(destination.read_text())["grading"]["status"] == "source_mutated"
