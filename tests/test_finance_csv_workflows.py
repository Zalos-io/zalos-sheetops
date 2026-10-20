"""Public CSV finance workflows preserve complete populations and exact evidence."""

import csv
import hashlib
from pathlib import Path

import pytest

from zalos_sheetops.errors import SheetOpsError


def csv_source(path: Path, headers: list[str], rows: list[list[str]]) -> Path:
    with path.open("w", encoding="utf-16", newline="") as stream:
        writer = csv.writer(stream, delimiter="|")
        writer.writerow(headers)
        writer.writerows(rows)
    return path


def test_reconcile_csv_checks_exact_tolerance_and_both_source_assignments(
    tmp_path: Path,
) -> None:
    from zalos_sheetops.finance import reconcile_csv

    left = csv_source(
        tmp_path / "bank.csv",
        [" Txn ID ", " Bank Amount ", "Memo"],
        [
            [" 0001 ", "10.005", "quoted\nrecord"],
            ["0002", "2.000", ""],
            ["0003", "-1.000", ""],
            [" bank-only ", "7", ""],
        ],
    )
    right = csv_source(
        tmp_path / "book.csv",
        ["Posting ID", "Book Amount"],
        [
            [" 0001 ", "10.000"],
            ["0002", "2.006"],
            ["0003", "-1.000"],
            ["ledger-only", "3"],
        ],
    )
    before = {"left": left.read_bytes(), "right": right.read_bytes()}
    analysis = reconcile_csv(
        left,
        right,
        left_key_columns=["Txn ID"],
        right_key_columns=["Posting ID"],
        left_amount_column="Bank Amount",
        right_amount_column="Book Amount",
        amount_tolerance="0.005",
        left_encoding="utf-16",
        right_encoding="utf-16",
        left_delimiter="|",
        right_delimiter="|",
    )
    assert analysis["status"] == "checked"
    report = analysis["report"]
    assert report["summary"] == {
        "matched": 2,
        "variance": 1,
        "leftUnmatched": 1,
        "rightUnmatched": 1,
    }
    assert report["variance"][0]["key"] == {"txn id": "0002"}
    assert report["variance"][0]["difference"] == "-0.006"
    assert analysis["checks"]["independentReconciliation"] is True
    assert analysis["checks"]["uniqueAssignments"] is True
    assert analysis["invalidRecordCount"] == 0
    for side, path in [("left", left), ("right", right)]:
        evidence = analysis["sources"][side]
        assert evidence["rowCount"] == 4 and evidence["invalidRecordCount"] == 0
        assert evidence["sha256Before"] == evidence["sha256After"]
        assert evidence["sha256Before"] == hashlib.sha256(before[side]).hexdigest()
        assert path.read_bytes() == before[side]


def test_reconcile_csv_counts_all_invalid_records_once_before_matching(
    tmp_path: Path,
) -> None:
    from zalos_sheetops.finance import reconcile_csv

    left = csv_source(
        tmp_path / "left.csv",
        ["ID", "Amount"],
        [[" ", "NaN"] for _ in range(25)] + [["good", "1"]],
    )
    right = csv_source(
        tmp_path / "right.csv",
        ["Ref", "Value"],
        [[" ", "Infinity"] for _ in range(24)] + [["good", "1"]],
    )
    analysis = reconcile_csv(
        left,
        right,
        left_key_columns=["id"],
        right_key_columns=["ref"],
        left_amount_column="amount",
        right_amount_column="value",
        amount_tolerance="0",
        left_encoding="utf-16",
        right_encoding="utf-16",
        left_delimiter="|",
        right_delimiter="|",
    )
    assert analysis["status"] == "blocked"
    assert analysis["report"] is None
    assert analysis["invalidRecordCount"] == 49
    assert analysis["sources"]["left"]["rowCount"] == 26
    assert analysis["sources"]["right"]["rowCount"] == 25
    assert analysis["checks"]["matchingSkipped"] is True
    assert len(analysis["rejections"]) == 2
    assert all(r["details"]["issuesTruncated"] for r in analysis["rejections"])
    assert all(r["details"]["sourceExcerpts"] for r in analysis["rejections"])


def test_reconcile_csv_rejects_duplicate_keys_in_unmatched_population(
    tmp_path: Path,
) -> None:
    from zalos_sheetops.finance import reconcile_csv

    left = csv_source(
        tmp_path / "left.csv",
        ["id", "amount"],
        [["a", "1"], ["orphan", "2"], ["orphan", "3"]],
    )
    right = csv_source(tmp_path / "right.csv", ["id", "amount"], [["a", "1"]])
    with pytest.raises(SheetOpsError) as caught:
        reconcile_csv(
            left,
            right,
            left_key_columns=["id"],
            right_key_columns=["id"],
            left_amount_column="amount",
            right_amount_column="amount",
            amount_tolerance="0",
            left_encoding="utf-16",
            right_encoding="utf-16",
            left_delimiter="|",
            right_delimiter="|",
        )
    assert caught.value.error_code == "duplicate_key"


def test_grouped_totals_keeps_source_and_individually_rounded_breaks_separate(
    tmp_path: Path,
) -> None:
    from zalos_sheetops.finance import check_grouped_totals_csv

    path = csv_source(
        tmp_path / "amounts.csv",
        ["Group", "Side", "Amount"],
        [
            ["X", "debit", "+0.005"],
            ["X", "debit", "0.005"],
            ["X", "credit", "0.010"],
            ["Y", "debit", "-2.5"],
            ["Y", "credit", "-2.50"],
        ],
    )
    before = path.read_bytes()
    analysis = check_grouped_totals_csv(
        path,
        group_column="Group",
        side_column="Side",
        amount_column="Amount",
        decimal_places={"X": 2, "Y": 0},
        left_label="debit",
        right_label="credit",
        rounding="half_even",
        source_tolerance="0",
        encoding="utf-16",
        delimiter="|",
    )
    assert analysis["status"] == "checked"
    groups = analysis["report"]["groups"]
    assert [g["group"] for g in groups] == ["X", "Y"]
    assert groups[0]["source"]["matches"] is True
    assert groups[0]["output"]["matches"] is False
    assert groups[0]["output"]["difference"] == "-0.01"
    assert groups[0]["validated"] is False
    assert groups[1]["output"]["leftTotal"] == "-2"
    assert groups[1]["output"]["matches"] is True
    assert analysis["sources"]["source"]["rowCount"] == 5
    assert analysis["checks"]["independentTotals"] is True
    assert path.read_bytes() == before


@pytest.mark.parametrize("side,group", [("other", "X"), ("debit", "unexpected")])
def test_grouped_totals_never_drops_unassigned_records(
    tmp_path: Path,
    side: str,
    group: str,
) -> None:
    from zalos_sheetops.finance import check_grouped_totals_csv

    path = csv_source(
        tmp_path / "amounts.csv",
        ["group", "side", "amount"],
        [["X", "debit", "1"], ["X", "credit", "1"], [group, side, "99"]],
    )
    with pytest.raises(SheetOpsError) as caught:
        check_grouped_totals_csv(
            path,
            group_column="group",
            side_column="side",
            amount_column="amount",
            decimal_places={"X": 2},
            left_label="debit",
            right_label="credit",
            encoding="utf-16",
            delimiter="|",
        )
    assert caught.value.error_code == "unassigned_record"


def test_reconciliation_guard_rejects_duplicate_assignments_from_matcher(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import zalos_sheetops.finance as finance

    left = csv_source(tmp_path / "left.csv", ["id", "amount"], [["a", "1"], ["b", "2"]])
    right = csv_source(
        tmp_path / "right.csv", ["id", "amount"], [["a", "1"], ["b", "2"]]
    )
    real_matcher = finance.reconcile_exact

    def broken_matcher(*args: object, **kwargs: object) -> dict:
        report = real_matcher(*args, **kwargs)
        report["matched"][1] = report["matched"][0]
        return report

    monkeypatch.setattr(finance, "reconcile_exact", broken_matcher)
    with pytest.raises(SheetOpsError) as caught:
        finance.reconcile_csv(
            left,
            right,
            left_key_columns=["id"],
            right_key_columns=["id"],
            left_amount_column="amount",
            right_amount_column="amount",
            amount_tolerance="0",
            left_encoding="utf-16",
            right_encoding="utf-16",
            left_delimiter="|",
            right_delimiter="|",
        )
    assert caught.value.error_code == "validation_mismatch"


def test_totals_guard_rejects_false_balance_even_with_unchanged_counts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import zalos_sheetops.finance as finance

    path = csv_source(
        tmp_path / "amounts.csv",
        ["group", "side", "amount"],
        [["X", "debit", "0.004"], ["X", "credit", "0.006"]],
    )
    real_check = finance.check_totals

    def broken_check(*args: object, **kwargs: object) -> dict:
        report = real_check(*args, **kwargs)
        report["output"]["matches"] = True
        report["validated"] = True
        return report

    monkeypatch.setattr(finance, "check_totals", broken_check)
    with pytest.raises(SheetOpsError) as caught:
        finance.check_grouped_totals_csv(
            path,
            group_column="group",
            side_column="side",
            amount_column="amount",
            decimal_places={"X": 2},
            left_label="debit",
            right_label="credit",
            source_tolerance="0.01",
            encoding="utf-16",
            delimiter="|",
        )
    assert caught.value.error_code == "validation_mismatch"


def test_finance_source_change_after_read_blocks_the_result(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import zalos_sheetops.finance as finance

    left = csv_source(tmp_path / "left.csv", ["id", "amount"], [["a", "1"]])
    right = csv_source(tmp_path / "right.csv", ["id", "amount"], [["a", "1"]])
    real_matcher = finance.reconcile_exact

    def changing_matcher(*args: object, **kwargs: object) -> dict:
        report = real_matcher(*args, **kwargs)
        left.write_bytes(left.read_bytes() + b"changed")
        return report

    monkeypatch.setattr(finance, "reconcile_exact", changing_matcher)
    with pytest.raises(SheetOpsError) as caught:
        finance.reconcile_csv(
            left,
            right,
            left_key_columns=["id"],
            right_key_columns=["id"],
            left_amount_column="amount",
            right_amount_column="amount",
            amount_tolerance="0",
            left_encoding="utf-16",
            right_encoding="utf-16",
            left_delimiter="|",
            right_delimiter="|",
        )
    assert caught.value.error_code == "source_changed"
