import json
from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner


def test_reconcile_command_returns_success_with_variance_and_unmatched_csvs(
    tmp_path: Path,
) -> None:
    left = tmp_path / "left.csv"
    right = tmp_path / "right.csv"
    left.write_text("reference,amount\nA,10.00\nB,20.00\n", encoding="utf-8")
    right.write_text("reference,amount\nA,10.01\nC,30.00\n", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "reconcile",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left)},
                    "right": {"path": str(right)},
                    "leftKeyColumns": ["reference"],
                    "rightKeyColumns": ["reference"],
                    "leftAmountColumn": "amount",
                    "rightAmountColumn": "amount",
                    "amountTolerance": "0.00",
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["ok"] is True
    assert payload["command"] == "reconcile"
    assert payload["data"]["summary"] == {
        "matched": 0,
        "leftUnmatched": 1,
        "rightUnmatched": 1,
        "variance": 1,
    }
    assert payload["data"]["variance"] == [
        {
            "key": {"reference": "A"},
            "leftAmount": "10.00",
            "rightAmount": "10.01",
            "difference": "-0.01",
        }
    ]


def test_reconcile_command_loads_xlsx_sources_with_read_table(tmp_path: Path) -> None:
    left = tmp_path / "left.xlsx"
    right = tmp_path / "right.xlsx"
    _write_workbook(left, "10.00")
    _write_workbook(right, "10.004")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "reconcile",
            "--json",
            json.dumps(
                {
                    "left": {"path": str(left), "sheet": "Data", "headerRow": 1},
                    "right": {"path": str(right), "sheet": "Data", "headerRow": 1},
                    "leftKeyColumns": ["reference"],
                    "rightKeyColumns": ["reference"],
                    "leftAmountColumn": "amount",
                    "rightAmountColumn": "amount",
                    "amountTolerance": "0.01",
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["data"]["summary"] == {
        "matched": 1,
        "leftUnmatched": 0,
        "rightUnmatched": 0,
        "variance": 0,
    }


def test_reconcile_command_rejects_invalid_payload_fields() -> None:
    runner = CliRunner()

    cases = [
        (
            {
                "right": {"path": "right.csv"},
                "leftKeyColumns": ["reference"],
                "rightKeyColumns": ["reference"],
                "leftAmountColumn": "amount",
                "rightAmountColumn": "amount",
            },
            "missing_required_field",
            {"field": "left"},
        ),
        (
            {
                "left": {"path": "left.csv"},
                "leftKeyColumns": ["reference"],
                "rightKeyColumns": ["reference"],
                "leftAmountColumn": "amount",
                "rightAmountColumn": "amount",
            },
            "missing_required_field",
            {"field": "right"},
        ),
        (
            {
                "left": {"path": "left.csv"},
                "right": {"path": "right.csv"},
                "leftKeyColumns": [],
                "rightKeyColumns": ["reference"],
                "leftAmountColumn": "amount",
                "rightAmountColumn": "amount",
            },
            "invalid_field",
            {"field": "leftKeyColumns", "expected": "non-empty list[string]"},
        ),
        (
            {
                "left": {"path": "left.csv"},
                "right": {"path": "right.csv"},
                "leftKeyColumns": ["reference"],
                "rightKeyColumns": [],
                "leftAmountColumn": "amount",
                "rightAmountColumn": "amount",
            },
            "invalid_field",
            {"field": "rightKeyColumns", "expected": "non-empty list[string]"},
        ),
        (
            {
                "left": {"path": "left.csv"},
                "right": {"path": "right.csv"},
                "leftKeyColumns": ["reference"],
                "rightKeyColumns": ["reference"],
                "rightAmountColumn": "amount",
            },
            "missing_required_field",
            {"field": "leftAmountColumn"},
        ),
        (
            {
                "left": {"path": "left.csv"},
                "right": {"path": "right.csv"},
                "leftKeyColumns": ["reference"],
                "rightKeyColumns": ["reference"],
                "leftAmountColumn": "amount",
                "rightAmountColumn": 1,
            },
            "invalid_field",
            {"field": "rightAmountColumn", "expected": "string"},
        ),
        (
            {
                "left": {"path": "left.csv"},
                "right": {"path": "right.csv"},
                "leftKeyColumns": ["reference"],
                "rightKeyColumns": ["reference"],
                "leftAmountColumn": "amount",
                "rightAmountColumn": "amount",
                "amountTolerance": "bad",
            },
            "invalid_field",
            {"field": "amountTolerance", "expected": "decimal string"},
        ),
    ]

    for payload, error_code, details in cases:
        result = runner.invoke(
            app,
            ["reconcile", "--json", json.dumps(payload), "--raw-output"],
        )

        assert result.exit_code == 1
        response = json.loads(result.output)
        assert response["ok"] is False
        assert response["command"] == "reconcile"
        assert response["errorCode"] == error_code
        assert response["details"] == details


def test_reconcile_command_is_available_in_schema() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    payload = json.loads(result.output)
    commands = {command["name"]: command for command in payload["data"]["commands"]}
    assert commands["reconcile"]["status"] == "available"


def _write_workbook(path: Path, amount: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet.append(["reference", "amount"])
    worksheet.append(["A", amount])
    workbook.save(path)
