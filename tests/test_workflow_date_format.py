import shutil
from pathlib import Path

import pytest
from openpyxl import load_workbook

from zalos_sheetops.benchmarks.workflows.cases import build_cases
from zalos_sheetops.benchmarks.workflows.grade import grade_case


@pytest.mark.parametrize(
    "number_format,expected_failure",
    [
        (r"yyyy\-mm\-dd", False),
        ('YYYY"-"MM"-"DD', False),
        ("mm/dd/yyyy", True),
    ],
)
def test_report_grading_compares_date_display(
    tmp_path: Path, number_format: str, expected_failure: bool
) -> None:
    case = next(c for c in build_cases(tmp_path / "cases", 1) if c["kind"] == "report")
    workspace = tmp_path / "output"
    workspace.mkdir()
    output = workspace / "output.xlsx"
    shutil.copy2(Path(case["golden"]) / "output.xlsx", output)
    workbook = load_workbook(output)
    workbook["Data"]["B2"].number_format = number_format
    workbook.save(output)
    failures = grade_case(case, workspace)["failures"]
    assert ("date_format:Data!B2" in failures) == expected_failure
