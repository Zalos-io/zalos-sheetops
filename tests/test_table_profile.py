"""A compact profile must retain complete-population and exact-decimal checks."""

from decimal import localcontext
from pathlib import Path

import pytest

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_csv_table import read_csv_table


def table(rows: list[dict[str, str]]) -> dict:
    return {
        "headers": ["reference", "amount", "note"],
        "rows": rows,
        "rowCount": len(rows),
        "returnedRowCount": len(rows),
        "truncated": False,
    }


def profile(data: dict, **options: object) -> dict:
    from zalos_sheetops.validation.table_profile import profile_table

    return profile_table(
        data, key_columns=["reference"], amount_columns=["amount"], **options
    )


def test_profile_reports_exact_totals_and_preserves_order_and_text(
    tmp_path: Path,
) -> None:
    path = tmp_path / "source.csv"
    path.write_text(
        "reference;amount;note\n"
        ' 0002;10000.12345678901234567890123456789;"café\nmemo"\n'
        "0001;-10000.00000000000000000000000000000;last\n",
        encoding="utf-16",
    )
    before = path.read_bytes()
    data = read_csv_table(path, encoding="utf-16", delimiter=";")
    with localcontext() as context:
        context.prec = 3
        result = profile(data)
    assert result["validated"] is True
    assert result["rowCount"] == 2
    assert result["columns"]["reference"] == {
        "first": " 0002",
        "last": "0001",
        "min": " 0002",
        "max": "0001",
        "blankCount": 0,
    }
    assert result["columns"]["note"]["first"] == "café\nmemo"
    assert result["totals"]["amount"] == "0.12345678901234567890123456789"
    assert result["checks"] == {
        "completePopulation": True,
        "rectangularRows": True,
        "nonblankTextKeys": ["reference"],
        "finiteDecimalAmounts": ["amount"],
    }
    assert path.read_bytes() == before


@pytest.mark.parametrize(
    "change",
    [
        {"truncated": True},
        {"rowCount": None},
        {"rowCount": 9},
        {"returnedRowCount": 9},
        {"rowCount": True},
    ],
)
def test_profile_rejects_incomplete_population(change: dict) -> None:
    data = table([{"reference": "0001", "amount": "1", "note": ""}])
    with pytest.raises(SheetOpsError) as error:
        profile({**data, **change})
    assert error.value.error_code == "incomplete_population"


def test_profile_checks_every_row_and_reports_all_invalid_records() -> None:
    data = table(
        [
            {"reference": "ok", "amount": "1", "note": ""},
            {"reference": "", "amount": "NaN", "note": ""},
            {"reference": "unmatched", "amount": "Infinity", "note": ""},
        ]
    )
    data["sourceRows"] = [2, 4, 7]
    with pytest.raises(SheetOpsError) as error:
        profile(data)
    assert error.value.error_code == "invalid_table_values"
    assert error.value.details["invalidRecordCount"] == 2
    assert error.value.details["issueCount"] == 3
    assert error.value.details["issues"][0]["sourceRow"] == 4


@pytest.mark.parametrize(
    "rows",
    [[["0001", "1"]], [{"reference": "0001", "amount": "1", "note": "", "extra": "0"}]],
)
def test_profile_rejects_nonrectangular_data(rows: list) -> None:
    with pytest.raises(SheetOpsError) as error:
        profile(table(rows))
    assert error.value.error_code == "table_width_mismatch"


def test_profile_accepts_string_table_rows() -> None:
    result = profile(
        {
            **table([]),
            "rows": [["0001", "1.20", ""]],
            "rowCount": 1,
            "returnedRowCount": 1,
        }
    )
    assert result["totals"] == {"amount": "1.20"}


@pytest.mark.parametrize("value", [1.2, True, None, "", "1e129", "9" * 129])
def test_profile_rejects_inexact_or_invalid_amounts(value: object) -> None:
    with pytest.raises(SheetOpsError) as error:
        profile(table([{"reference": "0001", "amount": value, "note": ""}]))
    assert error.value.error_code == "invalid_table_values"


def test_profile_validates_source_row_receipt() -> None:
    data = table([{"reference": "0001", "amount": "1", "note": ""}])
    with pytest.raises(SheetOpsError) as error:
        profile({**data, "sourceRows": []})
    assert error.value.error_code == "incomplete_population"


def test_empty_population_cannot_establish_amount_validation() -> None:
    with pytest.raises(SheetOpsError) as error:
        profile(table([]))
    assert error.value.error_code == "empty_population"


def test_unknown_validation_column_is_rejected() -> None:
    from zalos_sheetops.validation.table_profile import profile_table

    with pytest.raises(SheetOpsError) as error:
        profile_table(table([]), amount_columns=["missing"])
    assert error.value.error_code == "missing_column"


def test_invalid_counts_cover_population_after_issue_examples_are_full() -> None:
    data = table([{"reference": "", "amount": "NaN", "note": ""} for _ in range(1000)])
    with pytest.raises(SheetOpsError) as error:
        profile(data)
    assert error.value.details["invalidRecordCount"] == 1000
    assert error.value.details["issueCount"] == 2000
    assert len(error.value.details["issues"]) == 20
    assert error.value.details["issuesTruncated"] is True
