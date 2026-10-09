import json
from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner


def test_normalize_preview_command_returns_json_preview(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Amount", "Date"])
    sheet.append(["$1,234.50", "06/01/2026"])
    workbook.save(path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "normalize",
            "preview",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "sheet": "Data",
                    "headerStrategy": "detect",
                    "headerRow": None,
                    "columns": [
                        {"source": "Amount", "kind": "amount", "name": "amount"},
                        {"source": "Date", "kind": "date", "name": "date"},
                    ],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["ok"] is True
    assert payload["command"] == "normalize preview"
    assert (
        payload["data"]["previewRows"][0]["values"]["amount"]["normalized"] == "1234.50"
    )


def test_normalize_preview_rejects_boolean_header_row(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Amount"])
    sheet.append(["123.45"])
    workbook.save(path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "normalize",
            "preview",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "sheet": "Data",
                    "headerStrategy": "row",
                    "headerRow": True,
                    "columns": [
                        {"source": "Amount", "kind": "amount", "name": "amount"},
                    ],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    payload = json.loads(result.output)
    assert payload["ok"] is False
    assert payload["command"] == "normalize preview"
    assert payload["errorCode"] == "invalid_field"
    assert payload["details"] == {
        "field": "headerRow",
        "expected": "integer | null",
    }
