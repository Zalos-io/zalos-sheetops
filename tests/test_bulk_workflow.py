import csv
from pathlib import Path

from zalos_sheetops.benchmarks.workflows.cases import build_cases


def test_bulk_workflow_has_independent_control_totals(tmp_path: Path) -> None:
    cases = {case["id"]: case for case in build_cases(tmp_path, detail_rows=12)}
    case = cases["bulk_analysis"]
    with (Path(case["golden"]) / "output.csv").open() as stream:
        assert list(csv.reader(stream)) == [
            ["region", "count", "total", "first_id"],
            ["east", "2", "0.18", "000000000000000006"],
            ["north", "2", "0.14", "000000000000000004"],
            ["west", "2", "0.10", "000000000000000002"],
        ]
