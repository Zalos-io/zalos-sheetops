import pytest

from zalos_sheetops.diff.tabular import diff_tables
from zalos_sheetops.errors import SheetOpsError


def test_diff_tables_reports_added_removed_and_changed_rows() -> None:
    left = {
        "rows": [
            {"reference": "A", "amount": "10"},
            {"reference": "B", "amount": "20"},
        ]
    }
    right = {
        "rows": [
            {"reference": "A", "amount": "11"},
            {"reference": "C", "amount": "30"},
        ]
    }

    result = diff_tables(left, right, ["reference"], ["amount"], "0.00")

    assert result["same"] is False
    assert result["removed"] == [
        {"key": {"reference": "B"}, "left": {"reference": "B", "amount": "20"}}
    ]
    assert result["added"] == [
        {"key": {"reference": "C"}, "right": {"reference": "C", "amount": "30"}}
    ]
    assert result["changed"] == [
        {
            "key": {"reference": "A"},
            "changes": [{"column": "amount", "left": "10", "right": "11"}],
        }
    ]
    assert result["summary"] == {"added": 1, "removed": 1, "changed": 1}


def test_diff_tables_respects_decimal_amount_tolerance() -> None:
    left = {"rows": [{"reference": "A", "amount": "10.00"}]}
    right = {"rows": [{"reference": "A", "amount": "10.004"}]}

    result = diff_tables(left, right, ["reference"], ["amount"], "0.01")

    assert result == {
        "same": True,
        "added": [],
        "removed": [],
        "changed": [],
        "summary": {"added": 0, "removed": 0, "changed": 0},
    }


def test_diff_tables_treats_nonnumeric_mismatches_as_differences() -> None:
    left = {"rows": [{"reference": "A", "status": "paid"}]}
    right = {"rows": [{"reference": "A", "status": "pending"}]}

    result = diff_tables(left, right, ["reference"], None, "0.01")

    assert result["same"] is False
    assert result["changed"] == [
        {
            "key": {"reference": "A"},
            "changes": [{"column": "status", "left": "paid", "right": "pending"}],
        }
    ]


def test_diff_tables_rejects_duplicate_left_keys() -> None:
    left = {
        "rows": [
            {"reference": "A", "amount": "10"},
            {"reference": "A", "amount": "11"},
        ]
    }
    right = {"rows": [{"reference": "A", "amount": "11"}]}

    with pytest.raises(SheetOpsError) as error_info:
        diff_tables(left, right, ["reference"], ["amount"], "0.00")

    assert error_info.value.error_code == "duplicate_key"
    assert error_info.value.details == {"side": "left", "key": {"reference": "A"}}


def test_diff_tables_rejects_missing_key_columns() -> None:
    left = {"rows": [{"amount": "10"}]}
    right = {"rows": [{"reference": "A", "amount": "10"}]}

    with pytest.raises(SheetOpsError) as error_info:
        diff_tables(left, right, ["reference"], ["amount"], "0.00")

    assert error_info.value.error_code == "missing_key_column"
    assert error_info.value.details == {"side": "left", "columns": ["reference"]}


def test_diff_tables_rejects_missing_compare_columns() -> None:
    left = {"rows": [{"reference": "A"}]}
    right = {"rows": [{"reference": "A"}]}

    with pytest.raises(SheetOpsError) as error_info:
        diff_tables(left, right, ["reference"], ["amount"], "0.00")

    assert error_info.value.error_code == "missing_compare_column"
    assert error_info.value.details == {"side": "left", "columns": ["amount"]}


def test_diff_tables_rejects_empty_key_columns() -> None:
    table = {"rows": [{"reference": "A", "amount": "10"}]}

    with pytest.raises(SheetOpsError) as error_info:
        diff_tables(table, table, [], ["amount"], "0.00")

    assert error_info.value.error_code == "invalid_field"
    assert error_info.value.details == {
        "field": "key_columns",
        "expected": "non-empty list[string]",
    }


def test_diff_tables_rejects_empty_compare_columns() -> None:
    table = {"rows": [{"reference": "A", "amount": "10"}]}

    with pytest.raises(SheetOpsError) as error_info:
        diff_tables(table, table, ["reference"], [], "0.00")

    assert error_info.value.error_code == "invalid_field"
    assert error_info.value.details == {
        "field": "compare_columns",
        "expected": "non-empty list[string] | null",
    }
