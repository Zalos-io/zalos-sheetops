"""A source-verified CSV receipt retains all validation and original excerpts."""

import hashlib
import importlib
import json
from pathlib import Path

import pytest

from zalos_sheetops.errors import SheetOpsError


def reader() -> object:
    module = importlib.import_module("zalos_sheetops.io.read_csv_profile")
    return module.read_csv_profile


def test_complete_csv_profile_preserves_precision_multiline_and_source(
    tmp_path: Path,
) -> None:
    path = tmp_path / "source.csv"
    path.write_text(
        "Voucher ID;Net Balance;Note\n"
        '0002;0.12345678901234567890123456789012345;"memo\ncafé"\n'
        "0001;1;last\n",
        encoding="utf-16",
    )
    before = path.read_bytes()
    table, receipt = reader()(
        path,
        key_columns=["voucher id"],
        amount_columns=["net balance"],
        encoding="auto",
        delimiter=";",
    )
    assert receipt["validated"] is True and receipt["rowCount"] == 2
    assert receipt["totals"]["net balance"] == "1.12345678901234567890123456789012345"
    assert table["sourceRows"] == [2, 4] and table["truncated"] is False
    assert table["rows"][0]["note"] == "memo\ncafé"
    assert receipt["source"]["sha256"] == hashlib.sha256(before).hexdigest()
    assert receipt["source"]["headers"] == ["voucher id", "net balance", "note"]
    assert receipt["source"]["encoding"] == "utf-16"
    assert path.read_bytes() == before


def test_rejection_contains_original_source_excerpts_and_complete_counts(
    tmp_path: Path,
) -> None:
    path = tmp_path / "source.csv"
    path.write_text('document,value\n,NaN\n0002,"bad\namount"\n', encoding="utf-16")
    original = path.read_bytes()
    with pytest.raises(SheetOpsError) as caught:
        reader()(path, key_columns=["document"], amount_columns=["value"])
    error = caught.value
    assert error.error_code == "invalid_table_values"
    assert error.details["invalidRecordCount"] == 2
    assert error.details["issueCount"] == 3
    assert error.details["source"]["rowCount"] == 2
    assert error.details["source"]["sha256"] == hashlib.sha256(original).hexdigest()
    excerpts = error.details["sourceExcerpts"]
    assert [item["sourceRow"] for item in excerpts] == [2, 3]
    assert excerpts[0]["textPreview"].startswith(",NaN\n")
    assert excerpts[1]["textPreview"] == '0002,"bad\namount"\n'
    json.dumps(error.details, allow_nan=False)
    assert path.read_bytes() == original


def test_source_excerpts_are_bounded_without_losing_full_invalid_count(
    tmp_path: Path,
) -> None:
    path = tmp_path / "source.csv"
    path.write_text(
        "document,value\n" + "".join(f"000{i},{'bad' * 300}\n" for i in range(31)),
        encoding="utf-8",
    )
    with pytest.raises(SheetOpsError) as caught:
        reader()(path, key_columns=["document"], amount_columns=["value"])
    details = caught.value.details
    assert details["invalidRecordCount"] == details["issueCount"] == 31
    assert details["issuesTruncated"] is True
    assert len(details["sourceExcerpts"]) == 20
    assert all(len(row["textPreview"]) <= 240 for row in details["sourceExcerpts"])
    assert all(row["textTruncated"] for row in details["sourceExcerpts"])


def test_missing_column_retains_rejection_and_available_source_headers(
    tmp_path: Path,
) -> None:
    path = tmp_path / "source.csv"
    path.write_text("document,value\n0001,10\n", encoding="utf-8")
    with pytest.raises(SheetOpsError) as caught:
        reader()(path, key_columns=["document"], amount_columns=["amount"])
    assert caught.value.error_code == "missing_column"
    assert caught.value.details["columns"] == ["amount"]
    assert caught.value.details["source"]["headers"] == ["document", "value"]


@pytest.mark.parametrize("value", [None, "amount"])
def test_explicit_validation_columns_cannot_be_replaced_by_implicit_defaults(
    tmp_path: Path,
    value: object,
) -> None:
    path = tmp_path / "source.csv"
    path.write_text("document,value\n0001,10\n", encoding="utf-8")
    with pytest.raises(SheetOpsError) as caught:
        reader()(path, key_columns=["document"], amount_columns=value)
    assert caught.value.error_code == "invalid_field"


def test_source_change_during_validation_blocks_receipt(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "source.csv"
    path.write_text("document,value\n0001,10\n", encoding="utf-8")
    module = importlib.import_module("zalos_sheetops.io.read_csv_profile")
    original = module.profile_table

    def change_source(table: dict, **kwargs: object) -> dict:
        receipt = original(table, **kwargs)
        path.write_text("document,value\n0001,999\n", encoding="utf-8")
        return receipt

    monkeypatch.setattr(module, "profile_table", change_source)
    with pytest.raises(SheetOpsError) as caught:
        module.read_csv_profile(
            path, key_columns=["document"], amount_columns=["value"]
        )
    assert caught.value.error_code == "source_changed"


def test_missing_source_is_a_structured_rejection(tmp_path: Path) -> None:
    with pytest.raises(SheetOpsError) as caught:
        reader()(
            tmp_path / "missing.csv", key_columns=["document"], amount_columns=["value"]
        )
    assert caught.value.error_code == "file_not_found"


@pytest.mark.parametrize("limit", [1, 0, True])
def test_partial_source_requests_are_rejected_before_reading(
    tmp_path: Path,
    limit: object,
) -> None:
    with pytest.raises(SheetOpsError) as caught:
        reader()(
            tmp_path / "missing.csv",
            key_columns=["document"],
            amount_columns=["value"],
            max_rows=limit,
        )
    assert caught.value.error_code == "incomplete_population"


def test_changed_rejected_source_never_returns_stale_excerpts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    path = tmp_path / "source.csv"
    path.write_text("document,value\n0001,NaN\n", encoding="utf-8")
    module = importlib.import_module("zalos_sheetops.io.read_csv_profile")
    original = module.profile_table

    def change_rejected_source(table: dict, **kwargs: object) -> dict:
        try:
            return original(table, **kwargs)
        except SheetOpsError:
            path.write_text("", encoding="utf-8")
            raise

    monkeypatch.setattr(module, "profile_table", change_rejected_source)
    with pytest.raises(SheetOpsError) as caught:
        module.read_csv_profile(
            path, key_columns=["document"], amount_columns=["value"]
        )
    assert caught.value.error_code == "source_changed"
