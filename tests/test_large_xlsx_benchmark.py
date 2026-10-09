import json
from pathlib import Path

from zalos_sheetops.benchmarks.large_xlsx import (
    generate_large_xlsx_fixture,
    run_benchmark_plan,
)


def test_large_xlsx_benchmark_writes_inspect_and_validate_jsonl(
    tmp_path: Path,
) -> None:
    workbook_path = tmp_path / "large.xlsx"
    output_path = tmp_path / "results.jsonl"

    result = run_benchmark_plan(
        workbook_path=workbook_path,
        output_path=output_path,
        rows=5,
        columns=4,
        scenarios=["inspect", "validate"],
    )

    assert result == {
        "workbook": str(workbook_path),
        "output": str(output_path),
        "scenarioCount": 2,
    }
    records = [
        json.loads(line)
        for line in output_path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    assert [record["scenario"] for record in records] == ["inspect", "validate"]
    assert all(record["ok"] is True for record in records)
    assert all(record["rows"] == 5 for record in records)
    assert all(record["columns"] == 4 for record in records)
    assert all(record["elapsedSeconds"] >= 0 for record in records)
    assert all(record["peakTracedBytes"] >= 0 for record in records)
    assert records[0]["summary"] == {
        "engine": "xml",
        "sheetCount": 1,
        "formulaCount": 5,
    }
    assert records[1]["summary"] == {
        "valid": True,
        "issueCount": 0,
        "uniqueKeysChecked": {"Data": ["record_id"]},
    }


def test_large_xlsx_benchmark_covers_table_and_normalize_preview(
    tmp_path: Path,
) -> None:
    workbook_path = tmp_path / "large.xlsx"
    output_path = tmp_path / "results.jsonl"

    result = run_benchmark_plan(
        workbook_path=workbook_path,
        output_path=output_path,
        rows=5,
        columns=4,
        scenarios=["read-table", "normalize-preview"],
    )

    assert result["scenarioCount"] == 2
    records = [
        json.loads(line)
        for line in output_path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    assert [record["scenario"] for record in records] == [
        "read-table",
        "normalize-preview",
    ]
    assert all(record["ok"] is True for record in records)
    assert records[0]["summary"] == {
        "sheet": "Data",
        "rowCount": 5,
        "returnedRowCount": 5,
        "truncated": False,
    }
    assert records[1]["summary"] == {
        "rowCount": 5,
        "previewRowCount": 5,
        "issueCount": 0,
    }


def test_large_xlsx_benchmark_covers_paginated_cell_search(
    tmp_path: Path,
) -> None:
    workbook_path = tmp_path / "large.xlsx"
    output_path = tmp_path / "results.jsonl"

    result = run_benchmark_plan(
        workbook_path=workbook_path,
        output_path=output_path,
        rows=250,
        columns=4,
        scenarios=["cells-search"],
    )

    assert result["scenarioCount"] == 1
    record = json.loads(output_path.read_text(encoding="utf-8"))
    assert record["ok"] is True
    assert record["scenario"] == "cells-search"
    assert record["summary"]["returnedCount"] == 100
    assert record["summary"]["hasMore"] is True
    assert record["summary"]["stoppedEarly"] is False
    assert record["summary"]["scannedCellCount"] > 100


def test_generate_large_xlsx_fixture_is_deterministic(tmp_path: Path) -> None:
    workbook_path = tmp_path / "fixture.xlsx"

    metadata = generate_large_xlsx_fixture(workbook_path, rows=3, columns=5)

    assert metadata == {
        "path": str(workbook_path),
        "sheet": "Data",
        "rows": 3,
        "columns": 5,
        "headers": ["record_id", "amount", "status", "formula_1", "formula_2"],
    }
    assert workbook_path.exists()


def test_large_xlsx_benchmark_can_write_cpu_profiles(tmp_path: Path) -> None:
    workbook_path = tmp_path / "profiled.xlsx"
    output_path = tmp_path / "results.jsonl"
    profile_dir = tmp_path / "profiles"

    run_benchmark_plan(
        workbook_path=workbook_path,
        output_path=output_path,
        rows=2,
        columns=4,
        scenarios=["inspect"],
        cpu_profile_dir=profile_dir,
    )

    assert (profile_dir / "inspect.prof").exists()
