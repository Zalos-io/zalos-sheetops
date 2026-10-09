import argparse
import cProfile
import json
import sys
import time
import tracemalloc
from collections.abc import Callable
from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.inspect import inspect_workbook
from zalos_sheetops.io.read_table import read_table
from zalos_sheetops.io.search_cells import search_cells
from zalos_sheetops.normalize.preview import preview_normalization
from zalos_sheetops.validation.workbook import validate_workbook

DEFAULT_SCENARIOS = [
    "inspect",
    "validate",
    "read-table",
    "normalize-preview",
    "cells-search",
]
DEFAULT_ROWS = 100_000
DEFAULT_COLUMNS = 12
SHEET_NAME = "Data"
KEY_COLUMN = "record_id"


def generate_large_xlsx_fixture(
    path: Path,
    *,
    rows: int,
    columns: int,
) -> dict[str, object]:
    _require_positive_int(rows, "rows")
    _require_positive_int(columns, "columns")
    if columns < 4:
        raise ValueError("columns must be at least 4")

    path.parent.mkdir(parents=True, exist_ok=True)
    headers = _headers(columns)
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet(SHEET_NAME)
    sheet.append(headers)
    for row_number in range(1, rows + 1):
        sheet.append(_data_row(row_number, columns))
    workbook.save(path)

    return {
        "path": str(path),
        "sheet": SHEET_NAME,
        "rows": rows,
        "columns": columns,
        "headers": headers,
    }


def run_benchmark_plan(
    *,
    workbook_path: Path,
    output_path: Path,
    rows: int,
    columns: int,
    scenarios: list[str] | None = None,
    cpu_profile_dir: Path | None = None,
) -> dict[str, object]:
    scenario_names = _normalize_scenarios(scenarios)
    generate_large_xlsx_fixture(workbook_path, rows=rows, columns=columns)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if cpu_profile_dir is not None:
        cpu_profile_dir.mkdir(parents=True, exist_ok=True)

    records = [
        run_benchmark_scenario(
            scenario,
            workbook_path=workbook_path,
            rows=rows,
            columns=columns,
            cpu_profile_dir=cpu_profile_dir,
        )
        for scenario in scenario_names
    ]
    output_path.write_text(
        "".join(f"{json.dumps(record, sort_keys=True)}\n" for record in records),
        encoding="utf-8",
    )
    return {
        "workbook": str(workbook_path),
        "output": str(output_path),
        "scenarioCount": len(records),
    }


def run_benchmark_scenario(
    scenario: str,
    *,
    workbook_path: Path,
    rows: int,
    columns: int,
    cpu_profile_dir: Path | None = None,
) -> dict[str, object]:
    operation = _scenario_operation(scenario, workbook_path)
    profile_path = (
        cpu_profile_dir / f"{scenario}.prof" if cpu_profile_dir is not None else None
    )
    measurement = _measure(operation, profile_path)
    record = {
        "scenario": scenario,
        "ok": measurement["ok"],
        "workbook": str(workbook_path),
        "rows": rows,
        "columns": columns,
        "elapsedSeconds": measurement["elapsedSeconds"],
        "peakTracedBytes": measurement["peakTracedBytes"],
        "maxRssBytes": _max_rss_bytes(),
    }
    if measurement["ok"]:
        record["summary"] = _scenario_summary(scenario, measurement["data"])
    else:
        record["error"] = measurement["error"]
    return record


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    result = run_benchmark_plan(
        workbook_path=args.workbook,
        output_path=args.output,
        rows=args.rows,
        columns=args.columns,
        scenarios=_scenario_args(args.scenarios),
        cpu_profile_dir=args.cpu_profile_dir,
    )
    print(json.dumps(result, sort_keys=True))
    return 0


def _headers(columns: int) -> list[str]:
    base_headers = [KEY_COLUMN, "amount", "status"]
    formula_headers = [
        f"formula_{index}" for index in range(1, columns - len(base_headers) + 1)
    ]
    return base_headers + formula_headers


def _data_row(row_number: int, columns: int) -> list[object]:
    excel_row = row_number + 1
    formula_count = columns - 3
    return [
        f"REC-{row_number:07d}",
        row_number * 10,
        "open" if row_number % 2 else "closed",
        *[f"=B{excel_row}*{index}" for index in range(1, formula_count + 1)],
    ]


def _normalize_scenarios(scenarios: list[str] | None) -> list[str]:
    names = scenarios or DEFAULT_SCENARIOS
    invalid = [scenario for scenario in names if scenario not in DEFAULT_SCENARIOS]
    if not names or invalid:
        raise ValueError(
            "scenarios must be a non-empty list of supported scenarios",
        )
    return names


def _scenario_operation(
    scenario: str,
    workbook_path: Path,
) -> Callable[[], dict[str, object]]:
    if scenario == "inspect":
        return lambda: inspect_workbook(
            workbook_path,
            engine="xml",
            max_formula_cells=0,
        )
    if scenario == "validate":
        return lambda: validate_workbook(
            workbook_path,
            required_sheets=[SHEET_NAME],
            required_columns={SHEET_NAME: _headers(4)[:3]},
            unique_keys={SHEET_NAME: [KEY_COLUMN]},
        )
    if scenario == "read-table":
        return lambda: read_table(
            workbook_path,
            SHEET_NAME,
            "row",
            1,
            max_rows=100,
        )
    if scenario == "normalize-preview":
        return lambda: preview_normalization(
            workbook_path,
            SHEET_NAME,
            "row",
            1,
            [{"source": "amount", "kind": "amount", "name": "amount"}],
            max_rows=100,
        )
    if scenario == "cells-search":
        return lambda: search_cells(
            workbook_path,
            "open",
            fields=["values"],
            sheet_name=SHEET_NAME,
            limit=100,
        )
    raise ValueError(f"unknown scenario: {scenario}")


def _measure(
    operation: Callable[[], dict[str, object]],
    profile_path: Path | None,
) -> dict[str, object]:
    tracemalloc.start()
    started = time.perf_counter()
    profiler = cProfile.Profile() if profile_path is not None else None
    try:
        if profiler is not None:
            profiler.enable()
        data = operation()
        ok = True
        error = None
    except SheetOpsError as sheetops_error:
        data = {}
        ok = False
        error = {
            "errorCode": sheetops_error.error_code,
            "message": sheetops_error.message,
            "details": sheetops_error.details,
        }
    except Exception as unexpected_error:
        data = {}
        ok = False
        error = {
            "errorCode": unexpected_error.__class__.__name__,
            "message": str(unexpected_error),
            "details": {},
        }
    finally:
        if profiler is not None:
            profiler.disable()
            profiler.dump_stats(profile_path)
        elapsed = time.perf_counter() - started
        _, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()
    return {
        "ok": ok,
        "data": data,
        "error": error,
        "elapsedSeconds": round(elapsed, 6),
        "peakTracedBytes": peak,
    }


def _scenario_summary(
    scenario: str,
    data: dict[str, object],
) -> dict[str, object]:
    if scenario == "inspect":
        workbook = data["workbook"]
        if not isinstance(workbook, dict):
            return {}
        return {
            "engine": workbook.get("engine"),
            "sheetCount": workbook.get("sheetCount"),
            "formulaCount": workbook.get("formulaCount"),
        }
    if scenario == "validate":
        issues = data.get("issues", [])
        return {
            "valid": data.get("valid"),
            "issueCount": len(issues) if isinstance(issues, list) else 0,
            "uniqueKeysChecked": data.get("uniqueKeysChecked"),
        }
    if scenario == "read-table":
        return {
            "sheet": data.get("sheet"),
            "rowCount": data.get("rowCount"),
            "returnedRowCount": data.get("returnedRowCount"),
            "truncated": data.get("truncated"),
        }
    if scenario == "normalize-preview":
        preview_rows = data.get("previewRows", [])
        return {
            "rowCount": data.get("rowCount"),
            "previewRowCount": len(preview_rows)
            if isinstance(preview_rows, list)
            else 0,
            "issueCount": data.get("issueCount"),
        }
    if scenario == "cells-search":
        return {
            "returnedCount": data.get("returnedCount"),
            "hasMore": data.get("hasMore"),
            "stoppedEarly": data.get("stoppedEarly"),
            "scannedCellCount": data.get("scannedCellCount"),
        }
    return {}


def _max_rss_bytes() -> int | None:
    try:
        import resource
    except ImportError:
        return None

    max_rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    if sys.platform == "darwin":
        return int(max_rss)
    return int(max_rss) * 1024


def _scenario_args(value: str) -> list[str]:
    return [scenario.strip() for scenario in value.split(",") if scenario.strip()]


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Benchmark SheetOps large-XLSX inspection and validation.",
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path(".benchmarks/large-xlsx/workbook.xlsx"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(".benchmarks/large-xlsx/results.jsonl"),
    )
    parser.add_argument("--rows", type=int, default=DEFAULT_ROWS)
    parser.add_argument("--columns", type=int, default=DEFAULT_COLUMNS)
    parser.add_argument(
        "--scenarios",
        default=",".join(DEFAULT_SCENARIOS),
        help=(
            "Comma-separated scenarios: inspect,validate,read-table,"
            "normalize-preview,cells-search"
        ),
    )
    parser.add_argument(
        "--cpu-profile-dir",
        type=Path,
        default=None,
        help="Optional directory for cProfile .prof files.",
    )
    return parser.parse_args(argv)


def _require_positive_int(value: int, field_name: str) -> None:
    if isinstance(value, bool) or value < 1:
        raise ValueError(f"{field_name} must be a positive integer")


if __name__ == "__main__":
    raise SystemExit(main())
