import pytest

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.reconcile.exact import reconcile_exact


def test_reconcile_exact_classifies_matches_variances_and_unmatched() -> None:
    left = {
        "rows": [
            {"reference": "A", "amount": "10.00"},
            {"reference": "B", "amount": "20.00"},
            {"reference": "D", "amount": "40.00"},
        ]
    }
    right = {
        "rows": [
            {"reference": "A", "amount": "10.01"},
            {"reference": "C", "amount": "30.00"},
            {"reference": "D", "amount": "40.004"},
        ]
    }

    result = reconcile_exact(
        left,
        right,
        ["reference"],
        ["reference"],
        "amount",
        "amount",
        "0.005",
    )

    assert result["summary"] == {
        "matched": 1,
        "leftUnmatched": 1,
        "rightUnmatched": 1,
        "variance": 1,
    }
    assert result["matched"] == [
        {
            "key": {"reference": "D"},
            "left": {"reference": "D", "amount": "40.00"},
            "right": {"reference": "D", "amount": "40.004"},
        }
    ]
    assert result["variance"] == [
        {
            "key": {"reference": "A"},
            "leftAmount": "10.00",
            "rightAmount": "10.01",
            "difference": "-0.01",
        }
    ]
    assert result["leftUnmatched"] == [
        {"key": {"reference": "B"}, "left": {"reference": "B", "amount": "20.00"}}
    ]
    assert result["rightUnmatched"] == [
        {"key": {"reference": "C"}, "right": {"reference": "C", "amount": "30.00"}}
    ]


def test_reconcile_exact_matches_composite_keys_with_different_column_names() -> None:
    left = {"rows": [{"batch": "June", "reference": "A", "amount": "10.00"}]}
    right = {"rows": [{"period": "June", "id": "A", "total": "10.00"}]}

    result = reconcile_exact(
        left,
        right,
        ["batch", "reference"],
        ["period", "id"],
        "amount",
        "total",
        "0.00",
    )

    assert result["summary"] == {
        "matched": 1,
        "leftUnmatched": 0,
        "rightUnmatched": 0,
        "variance": 0,
    }
    assert result["matched"] == [
        {
            "key": {"batch": "June", "reference": "A"},
            "left": {"batch": "June", "reference": "A", "amount": "10.00"},
            "right": {"period": "June", "id": "A", "total": "10.00"},
        }
    ]


def test_reconcile_exact_rejects_empty_key_columns() -> None:
    table = {"rows": [{"reference": "A", "amount": "10.00"}]}

    with pytest.raises(SheetOpsError) as error_info:
        reconcile_exact(table, table, [], ["reference"], "amount", "amount", "0.00")

    assert error_info.value.error_code == "invalid_field"
    assert error_info.value.details == {
        "field": "left_key_columns",
        "expected": "non-empty list[string]",
    }


def test_reconcile_exact_rejects_invalid_decimal_tolerance() -> None:
    table = {"rows": [{"reference": "A", "amount": "10.00"}]}

    with pytest.raises(SheetOpsError) as error_info:
        reconcile_exact(
            table,
            table,
            ["reference"],
            ["reference"],
            "amount",
            "amount",
            "not-a-decimal",
        )

    assert error_info.value.error_code == "invalid_field"
    assert error_info.value.details == {
        "field": "amount_tolerance",
        "expected": "decimal string",
    }
