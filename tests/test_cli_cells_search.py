import json
from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner


def test_cells_search_command_returns_paginated_json_contract(
    tmp_path: Path,
) -> None:
    path = tmp_path / "search.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "Needle one"
    sheet["A2"] = "Needle two"
    workbook.save(path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "cells",
            "search",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "query": "needle",
                    "fields": ["values"],
                    "limit": 1,
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["ok"] is True
    assert payload["command"] == "cells search"
    assert payload["warnings"] == []
    assert payload["artifacts"] == []
    assert payload["data"]["returnedCount"] == 1
    assert payload["data"]["hasMore"] is True
    assert payload["data"]["nextOffset"] == 1
    assert payload["data"]["matches"][0] == {
        "sheet": "Data",
        "cell": "A1",
        "matchedFields": ["values"],
        "value": "Needle one",
    }


def test_cells_search_command_returns_stable_invalid_regex_error(
    tmp_path: Path,
) -> None:
    path = tmp_path / "search.xlsx"
    workbook = Workbook()
    workbook.save(path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "cells",
            "search",
            "--json",
            json.dumps({"path": str(path), "query": "(", "regex": True}),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    assert json.loads(result.output) == {
        "ok": False,
        "command": "cells search",
        "errorCode": "invalid_regex",
        "message": "Search pattern is not a valid regular expression",
        "details": {"query": "("},
        "warnings": [],
        "artifacts": [],
    }
