"""Read-only finance contracts beyond happy-path reconciliation."""

from decimal import Decimal, localcontext
from pathlib import Path

import pytest
from openpyxl import Workbook

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_table import read_table
from zalos_sheetops.reconcile.exact import reconcile_exact


def reconcile(left: dict, right: dict) -> dict:
    return reconcile_exact(
        left, right, ["reference"], ["reference"], "amount", "amount", "0"
    )


def test_table_read_rejects_populated_cells_without_headers(tmp_path: Path) -> None:
    path = tmp_path / "unheaded-amount.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Ledger"
    sheet.append(["reference", "amount"])
    sheet.append(["0001", "10.00", "999.00"])
    workbook.save(path)

    with pytest.raises(SheetOpsError) as error:
        read_table(path, "Ledger", "row", 1)

    assert error.value.error_code == "table_width_mismatch"
    assert error.value.details == {"sheet": "Ledger", "row": 2, "headerColumns": 2}


@pytest.mark.parametrize("side", ["left", "right"])
def test_reconcile_rejects_partial_population(side: str) -> None:
    complete = {"rows": [{"reference": "0001", "amount": "10.00"}]}
    preview = {**complete, "truncated": True, "rowCount": None}

    with pytest.raises(SheetOpsError) as error:
        reconcile(
            preview if side == "left" else complete,
            preview if side == "right" else complete,
        )

    assert error.value.error_code == "incomplete_population"
    assert error.value.details == {"side": side}


@pytest.mark.parametrize("key", [None, "", "   "])
def test_reconcile_rejects_blank_reference(key: object) -> None:
    table = {"rows": [{"reference": key, "amount": "10.00"}]}

    with pytest.raises(SheetOpsError) as error:
        reconcile(table, table)

    assert error.value.error_code == "invalid_key_value"
    assert error.value.details == {"side": "left", "column": "reference", "row": 1}


@pytest.mark.parametrize("value", ["NaN", "Infinity", "", None])
def test_reconcile_validates_unmatched_amounts(value: object) -> None:
    left = {"rows": [{"reference": "unmatched", "amount": value}]}
    right = {"rows": [{"reference": "other", "amount": "10.00"}]}

    with pytest.raises(SheetOpsError) as error:
        reconcile(left, right)

    assert error.value.error_code == "invalid_field"
    assert error.value.details["field"] == "leftAmountColumn"


def test_reconcile_difference_is_independent_of_decimal_context() -> None:
    left = {"rows": [{"reference": "0001", "amount": "10000.12345"}]}
    right = {"rows": [{"reference": "0001", "amount": "10000.00000"}]}
    with localcontext() as context:
        context.prec = 3
        result = reconcile(left, right)

    assert result["variance"][0]["difference"] == "0.12345"


def test_reconcile_preserves_distinct_keys_without_trimming() -> None:
    left = {"rows": [{"reference": " 0001", "amount": "10.00"}]}
    right = {"rows": [{"reference": "0001", "amount": "10.00"}]}
    assert reconcile(left, right)["summary"] == {
        "matched": 0,
        "variance": 0,
        "leftUnmatched": 1,
        "rightUnmatched": 1,
    }


@pytest.mark.parametrize("amount", ["1e129", "1e-129", "9" * 129])
def test_reconcile_rejects_amounts_outside_exact_precision_contract(
    amount: str,
) -> None:
    table = {"rows": [{"reference": "0001", "amount": amount}]}
    with pytest.raises(SheetOpsError) as error:
        reconcile(table, table)
    assert error.value.error_code == "invalid_field"


def test_reconcile_cli_rejects_invalid_unmatched_csv_amount(tmp_path: Path) -> None:
    import json

    from zalos_sheetops.cli import app
    from zalos_sheetops.testing import CliRunner

    left, right = tmp_path / "left.csv", tmp_path / "right.csv"
    left.write_text("reference,amount\n0001,10.00\n0002,NaN\n")
    right.write_text("reference,amount\n0001,10.00\n")
    payload = {
        "left": {"path": str(left)},
        "right": {"path": str(right)},
        "leftKeyColumns": ["reference"],
        "rightKeyColumns": ["reference"],
        "leftAmountColumn": "amount",
        "rightAmountColumn": "amount",
        "amountTolerance": "0.01",
    }
    result = CliRunner().invoke(app, ["reconcile", "--json", json.dumps(payload)])
    response = json.loads(result.output)
    assert result.exit_code == 1
    assert response["ok"] is False
    assert response["errorCode"] == "invalid_field"
    assert response["details"]["field"] == "leftAmountColumn"


@pytest.mark.parametrize(
    "key",
    [float("nan"), float("inf"), float("-inf"), Decimal("NaN"), Decimal("Infinity")],
)
def test_reconcile_rejects_nonfinite_numeric_reference(key: float | Decimal) -> None:
    table = {"rows": [{"reference": key, "amount": "10.00"}]}
    with pytest.raises(SheetOpsError) as error:
        reconcile(table, table)
    assert error.value.error_code == "invalid_key_value"
