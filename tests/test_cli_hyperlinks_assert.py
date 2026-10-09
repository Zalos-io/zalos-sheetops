import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.hyperlink import Hyperlink

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner

UPSTREAM_HYPERLINK_FIXTURE = Path("tests/fixtures/upstream/xlsxwriter/hyperlink01.xlsx")


def _write_hyperlink_fixture(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Links"
    sheet["B2"] = "Customer portal"
    sheet["B2"].hyperlink = Hyperlink(
        ref="B2",
        target="https://example.com/customer?id=42",
        tooltip="Open customer portal",
        display="Customer portal",
    )
    workbook.save(path)


def test_hyperlinks_assert_command_accepts_target_display_and_tooltip_checks(
    tmp_path: Path,
) -> None:
    path = tmp_path / "links.xlsx"
    _write_hyperlink_fixture(path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "hyperlinks",
            "assert",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "sheet": "Links",
                    "checks": [
                        {
                            "cell": "B2",
                            "target": "https://example.com/customer?id=42",
                            "display": "Customer portal",
                            "tooltip": "Open customer portal",
                        }
                    ],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    assert json.loads(result.output) == {
        "ok": True,
        "command": "hyperlinks assert",
        "data": {
            "valid": True,
            "sheet": "Links",
            "checkedCount": 1,
            "failureCount": 0,
            "failures": [],
        },
        "warnings": [],
        "artifacts": [],
    }


def test_hyperlinks_assert_command_accepts_upstream_xlsxwriter_fixture() -> None:
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "hyperlinks",
            "assert",
            "--json",
            json.dumps(
                {
                    "path": str(UPSTREAM_HYPERLINK_FIXTURE),
                    "sheet": "Sheet1",
                    "checks": [
                        {
                            "cell": "A1",
                            "target": "http://www.perl.org/",
                            "display": "http://www.perl.org/",
                        }
                    ],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["ok"] is True
    assert payload["data"]["failureCount"] == 0


def test_hyperlinks_assert_command_reports_target_mismatch(
    tmp_path: Path,
) -> None:
    path = tmp_path / "links.xlsx"
    _write_hyperlink_fixture(path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "hyperlinks",
            "assert",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "sheet": "Links",
                    "checks": [{"cell": "B2", "target": "https://example.com/wrong"}],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    payload = json.loads(result.output)
    assert payload["ok"] is False
    assert payload["command"] == "hyperlinks assert"
    assert payload["errorCode"] == "hyperlink_assertion_failed"
    assert payload["message"] == "Hyperlink assertions failed"
    assert payload["details"] == {
        "valid": False,
        "sheet": "Links",
        "checkedCount": 1,
        "failureCount": 1,
        "failures": [
            {
                "cell": "B2",
                "code": "target_mismatch",
                "expected": "https://example.com/wrong",
                "actual": "https://example.com/customer?id=42",
            }
        ],
    }


def test_hyperlinks_assert_command_rejects_checks_without_expectation(
    tmp_path: Path,
) -> None:
    path = tmp_path / "links.xlsx"
    _write_hyperlink_fixture(path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "hyperlinks",
            "assert",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "sheet": "Links",
                    "checks": [{"cell": "B2"}],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    assert json.loads(result.output) == {
        "ok": False,
        "command": "hyperlinks assert",
        "errorCode": "invalid_field",
        "message": "Field has invalid type",
        "details": {
            "field": "checks[0]",
            "expected": (
                "hyperlinkCheck with cell and at least one of target, display, tooltip"
            ),
        },
        "warnings": [],
        "artifacts": [],
    }
