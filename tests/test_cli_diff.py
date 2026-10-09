import json
from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner


def test_diff_command_returns_differences_found_for_changed_csvs(
    tmp_path: Path,
) -> None:
    left = tmp_path / "left.csv"
    right = tmp_path / "right.csv"
    left.write_text("reference,amount\nA,10\n", encoding="utf-8")
    right.write_text("reference,amount\nA,11\n", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "diff",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left)},
                    "right": {"path": str(right)},
                    "keyColumns": ["reference"],
                    "compareColumns": ["amount"],
                    "amountTolerance": "0.00",
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    payload = json.loads(result.output)
    assert payload["ok"] is False
    assert payload["command"] == "diff"
    assert payload["errorCode"] == "differences_found"
    assert payload["details"]["summary"] == {"added": 0, "removed": 0, "changed": 1}


def test_diff_command_returns_success_for_matching_csvs(tmp_path: Path) -> None:
    left = tmp_path / "left.csv"
    right = tmp_path / "right.csv"
    left.write_text("reference,amount\nA,10.00\n", encoding="utf-8")
    right.write_text("reference,amount\nA,10.004\n", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "diff",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left)},
                    "right": {"path": str(right)},
                    "keyColumns": ["reference"],
                    "compareColumns": ["amount"],
                    "amountTolerance": "0.01",
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload == {
        "ok": True,
        "command": "diff",
        "data": {
            "same": True,
            "added": [],
            "removed": [],
            "changed": [],
            "summary": {"added": 0, "removed": 0, "changed": 0},
        },
        "warnings": [],
        "artifacts": [],
    }


def test_diff_command_loads_xlsx_sources_with_read_table(tmp_path: Path) -> None:
    left = tmp_path / "left.xlsx"
    right = tmp_path / "right.xlsx"
    _write_workbook(left, "10")
    _write_workbook(right, "11")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "diff",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left), "sheet": "Data", "headerRow": 1},
                    "right": {"path": str(right), "sheet": "Data", "headerRow": 1},
                    "keyColumns": ["reference"],
                    "compareColumns": ["amount"],
                    "amountTolerance": "0.00",
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    payload = json.loads(result.output)
    assert payload["errorCode"] == "differences_found"
    assert payload["details"]["changed"] == [
        {
            "key": {"reference": "A"},
            "changes": [{"column": "amount", "left": "10", "right": "11"}],
        }
    ]


def test_diff_command_rejects_invalid_payload_fields() -> None:
    runner = CliRunner()

    cases = [
        (
            {"right": {"path": "right.csv"}, "keyColumns": ["reference"]},
            "missing_required_field",
            {"field": "left"},
        ),
        (
            {"left": {"path": "left.csv"}, "keyColumns": ["reference"]},
            "missing_required_field",
            {"field": "right"},
        ),
        (
            {"left": {"path": "left.csv"}, "right": {"path": "right.csv"}},
            "missing_required_field",
            {"field": "keyColumns"},
        ),
        (
            {
                "left": {"path": "left.csv"},
                "right": {"path": "right.csv"},
                "keyColumns": ["reference"],
                "compareColumns": "amount",
            },
            "invalid_field",
            {"field": "compareColumns", "expected": "list[string] | null"},
        ),
        (
            {
                "left": {"path": "left.csv"},
                "right": {"path": "right.csv"},
                "keyColumns": ["reference"],
                "amountTolerance": 0,
            },
            "invalid_field",
            {"field": "amountTolerance", "expected": "decimal string"},
        ),
    ]

    for payload, error_code, details in cases:
        result = runner.invoke(
            app,
            ["diff", "--json", json.dumps(payload), "--raw-output"],
        )

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["ok"] is False
        assert response["command"] == "diff"
        assert response["errorCode"] == error_code
        assert response["details"] == details


def test_diff_command_rejects_empty_key_columns(tmp_path: Path) -> None:
    left = tmp_path / "left.csv"
    right = tmp_path / "right.csv"
    left.write_text("reference,amount\nA,10\n", encoding="utf-8")
    right.write_text("reference,amount\nA,10\n", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "diff",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left)},
                    "right": {"path": str(right)},
                    "keyColumns": [],
                    "compareColumns": ["amount"],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    response = json.loads(result.output)
    assert response["ok"] is False
    assert response["command"] == "diff"
    assert response["errorCode"] == "invalid_field"
    assert response["details"] == {
        "field": "keyColumns",
        "expected": "non-empty list[string]",
    }


def test_diff_command_rejects_empty_compare_columns(tmp_path: Path) -> None:
    left = tmp_path / "left.csv"
    right = tmp_path / "right.csv"
    left.write_text("reference,amount\nA,10\n", encoding="utf-8")
    right.write_text("reference,amount\nA,10\n", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "diff",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left)},
                    "right": {"path": str(right)},
                    "keyColumns": ["reference"],
                    "compareColumns": [],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    response = json.loads(result.output)
    assert response["ok"] is False
    assert response["command"] == "diff"
    assert response["errorCode"] == "invalid_field"
    assert response["details"] == {
        "field": "compareColumns",
        "expected": "non-empty list[string] | null",
    }


def test_diff_command_surfaces_duplicate_key_errors_as_json(tmp_path: Path) -> None:
    left = tmp_path / "left.csv"
    right = tmp_path / "right.csv"
    left.write_text("reference,amount\nA,10\nA,11\n", encoding="utf-8")
    right.write_text("reference,amount\nA,11\n", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "diff",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left)},
                    "right": {"path": str(right)},
                    "keyColumns": ["reference"],
                    "compareColumns": ["amount"],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    response = json.loads(result.output)
    assert response["ok"] is False
    assert response["command"] == "diff"
    assert response["errorCode"] == "duplicate_key"
    assert response["details"] == {"side": "left", "key": {"reference": "A"}}


def test_diff_command_surfaces_missing_key_column_errors_as_json(
    tmp_path: Path,
) -> None:
    left = tmp_path / "left.csv"
    right = tmp_path / "right.csv"
    left.write_text("amount\n10\n", encoding="utf-8")
    right.write_text("amount\n10\n", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "diff",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left)},
                    "right": {"path": str(right)},
                    "keyColumns": ["reference"],
                    "compareColumns": ["amount"],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    response = json.loads(result.output)
    assert response == {
        "ok": False,
        "command": "diff",
        "errorCode": "missing_key_column",
        "message": "Required key column is missing from a row",
        "details": {"side": "left", "columns": ["reference"]},
        "warnings": [],
        "artifacts": [],
    }


def test_diff_command_surfaces_missing_compare_column_errors_as_json(
    tmp_path: Path,
) -> None:
    left = tmp_path / "left.csv"
    right = tmp_path / "right.csv"
    left.write_text("reference\nA\n", encoding="utf-8")
    right.write_text("reference\nA\n", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "diff",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left)},
                    "right": {"path": str(right)},
                    "keyColumns": ["reference"],
                    "compareColumns": ["amount"],
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    response = json.loads(result.output)
    assert response["ok"] is False
    assert response["command"] == "diff"
    assert response["errorCode"] == "missing_compare_column"
    assert response["details"] == {"side": "left", "columns": ["amount"]}


def test_schema_marks_diff_as_available() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    commands = {command["name"]: command for command in payload["data"]["commands"]}
    assert commands["diff"]["status"] == "available"


def _write_workbook(path: Path, amount: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["reference", "amount"])
    sheet.append(["A", amount])
    workbook.save(path)
