"""Read-only CSV finance workflows with explicit populations and postconditions."""

import hashlib
from decimal import ROUND_HALF_EVEN, ROUND_HALF_UP, Context, Decimal, localcontext
from pathlib import Path

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_csv_profile import read_csv_profile
from zalos_sheetops.normalize.text import canonical_header
from zalos_sheetops.reconcile.exact import reconcile_exact
from zalos_sheetops.validation.totals import check_totals


def _require(condition: bool, field: str) -> None:
    if not condition:
        raise SheetOpsError(
            "validation_mismatch", "Finance postcondition failed", {"field": field}
        )


def _column(value: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise SheetOpsError("invalid_field", "Supply a nonblank column name", {})
    return canonical_header(value)


def _keys(values: list[str]) -> list[str]:
    if not isinstance(values, list) or not values:
        raise SheetOpsError("invalid_field", "Supply a nonempty key column list", {})
    result = [_column(value) for value in values]
    if len(set(result)) != len(result):
        raise SheetOpsError("invalid_field", "Key columns must be distinct", {})
    return result


def _digest(path: Path) -> str:
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        raise SheetOpsError(
            "source_read_failed", "Cannot verify CSV source", {"path": str(path)}
        ) from error


def _read(
    path: Path, side: str, keys: list[str], amount: str, encoding: str, delimiter: str
) -> tuple[dict | None, dict, dict | None]:
    rejection = None
    try:
        table, profile = read_csv_profile(
            Path(path),
            key_columns=keys,
            amount_columns=[amount],
            encoding=encoding,
            delimiter=delimiter,
            max_rows=None,
        )
        source, invalid = profile["source"], 0
    except SheetOpsError as error:
        if error.error_code != "invalid_table_values":
            raise
        table = None
        source, invalid = error.details["source"], error.details["invalidRecordCount"]
        rejection = {
            "side": side,
            "errorCode": error.error_code,
            "details": error.details,
        }
    evidence = {
        "path": source["path"],
        "rowCount": source["rowCount"],
        "invalidRecordCount": invalid,
        "sha256Before": source["sha256"],
        "sha256After": source["sha256"],
        "headers": source["headers"],
        "encoding": source["encoding"],
        "delimiter": source["delimiter"],
        "headerRow": source["headerRow"],
    }
    return table, evidence, rejection


def _finish(
    sources: dict, rejections: list[dict], report: dict | None, checks: dict
) -> dict:
    for evidence in sources.values():
        after = _digest(Path(evidence["path"]))
        if after != evidence["sha256Before"]:
            raise SheetOpsError(
                "source_changed",
                "CSV source changed during finance checks",
                {"path": evidence["path"]},
            )
        evidence["sha256After"] = after
    return {
        "status": "blocked" if rejections else "checked",
        "report": report,
        "sources": sources,
        "rejections": rejections,
        "invalidRecordCount": sum(s["invalidRecordCount"] for s in sources.values()),
        "checks": {"completePopulation": True, "sourceUnchanged": True, **checks},
    }


def reconcile_csv(
    left_path: Path,
    right_path: Path,
    *,
    left_key_columns: list[str],
    right_key_columns: list[str],
    left_amount_column: str,
    right_amount_column: str,
    amount_tolerance: str,
    left_encoding: str = "auto",
    right_encoding: str = "auto",
    left_delimiter: str = ",",
    right_delimiter: str = ",",
) -> dict:
    """Profile both entire sources, then reconcile unique exact string keys.

    Column names are normalized like reader headers; key VALUES stay exact.
    Returns status='checked', report (the reconcile_exact dict), sources keyed
    left/right, rejections=[], invalidRecordCount=0 and checks. Each source has
    path, rowCount, invalidRecordCount, sha256Before/After, headers, encoding,
    delimiter, headerRow. Complete record-level rejections on either side return
    status='blocked', report=None, full invalidRecordCount and bounded rejections
    retaining sourceExcerpts. Both sides are scanned before that decision.

    Structural/schema/duplicate-key/source-change errors raise SheetOpsError;
    they never become successful empty results. Independent classification,
    amount-difference, unique-assignment and both-side coverage checks verify the
    matcher result. Mechanical checks do not approve accounting or posting.
    No files are written. Save only requested fields; do not print whole tables.
    """
    left_keys, right_keys = _keys(left_key_columns), _keys(right_key_columns)
    left_amount, right_amount = (
        _column(left_amount_column),
        _column(right_amount_column),
    )
    tables, sources, rejections = {}, {}, []
    for side, path, keys, amount, encoding, delimiter in [
        ("left", left_path, left_keys, left_amount, left_encoding, left_delimiter),
        (
            "right",
            right_path,
            right_keys,
            right_amount,
            right_encoding,
            right_delimiter,
        ),
    ]:
        table, source, rejection = _read(path, side, keys, amount, encoding, delimiter)
        tables[side], sources[side] = table, source
        if rejection is not None:
            rejections.append(rejection)
    if rejections:
        return _finish(sources, rejections, None, {"matchingSkipped": True})
    report = reconcile_exact(
        tables["left"],
        tables["right"],
        left_keys,
        right_keys,
        left_amount,
        right_amount,
        amount_tolerance,
    )
    _verify_reconciliation(
        tables["left"],
        tables["right"],
        report,
        left_keys,
        right_keys,
        left_amount,
        right_amount,
        amount_tolerance,
    )
    return _finish(
        sources,
        [],
        report,
        {
            "independentReconciliation": True,
            "uniqueAssignments": True,
        },
    )


def _verify_reconciliation(
    left: dict,
    right: dict,
    report: dict,
    left_keys: list[str],
    right_keys: list[str],
    left_amount: str,
    right_amount: str,
    tolerance: str,
) -> None:
    lhs = {tuple(row[c] for c in left_keys): row for row in left["rows"]}
    rhs = {tuple(row[c] for c in right_keys): row for row in right["rows"]}
    _require(
        len(lhs) == left["rowCount"] and len(rhs) == right["rowCount"], "uniqueKeys"
    )
    expected = {}
    with localcontext(Context(prec=400)):
        for key in lhs.keys() | rhs.keys():
            if key not in rhs:
                expected[key] = "leftUnmatched"
            elif key not in lhs:
                expected[key] = "rightUnmatched"
            else:
                difference = Decimal(lhs[key][left_amount]) - Decimal(
                    rhs[key][right_amount]
                )
                expected[key] = (
                    "matched" if abs(difference) <= Decimal(tolerance) else "variance"
                )
        seen = set()
        for kind in ("matched", "variance", "leftUnmatched", "rightUnmatched"):
            count = 0
            for item in report[kind]:
                columns = right_keys if kind == "rightUnmatched" else left_keys
                _require(set(item["key"]) == set(columns), "keyColumns")
                key = tuple(item["key"][c] for c in columns)
                _require(key not in seen and expected.get(key) == kind, "assignments")
                seen.add(key)
                count += 1
                _verify_match_values(
                    item, kind, key, lhs, rhs, left_amount, right_amount
                )
            _require(
                type(report["summary"][kind]) is int
                and report["summary"][kind] == count,
                "summary." + kind,
            )
        _require(seen == set(expected), "completeAssignments")


def _verify_match_values(
    item: dict,
    kind: str,
    key: tuple,
    lhs: dict,
    rhs: dict,
    left_amount: str,
    right_amount: str,
) -> None:
    if kind in {"matched", "leftUnmatched"}:
        _require(item["left"] == lhs[key], "leftSourceRecord")
    if kind in {"matched", "rightUnmatched"}:
        _require(item["right"] == rhs[key], "rightSourceRecord")
    if kind == "variance":
        left, right = Decimal(lhs[key][left_amount]), Decimal(rhs[key][right_amount])
        _require(Decimal(item["leftAmount"]) == left, "variance.leftAmount")
        _require(Decimal(item["rightAmount"]) == right, "variance.rightAmount")
        _require(Decimal(item["difference"]) == left - right, "variance.difference")


def check_grouped_totals_csv(
    path: Path,
    *,
    group_column: str,
    side_column: str,
    amount_column: str,
    decimal_places: dict[str, int],
    left_label: str,
    right_label: str,
    rounding: str = "half_up",
    source_tolerance: str = "0",
    encoding: str = "auto",
    delimiter: str = ",",
) -> dict:
    """Profile all records and compare source/per-line-rounded sums per group.

    All grouping values must occur in decimal_places, with both declared side
    labels present in every group. No record is dropped or netted across groups.
    Returns the same status/sources/rejections/invalidRecordCount/checks envelope
    as reconcile_csv, with one source named 'source'. On checked input, report
    is {'groups': [{'group': <source string>, **check_totals_report}, ...]}, sorted
    by group. A legitimate arithmetic break has group.validated=False and must
    remain visible; status='checked' describes execution, not balanced books.
    Independent decimal recomputation verifies counts, totals, differences and
    match flags, including source tolerance and zero output tolerance. No writes.
    """
    group, side, amount = map(_column, (group_column, side_column, amount_column))
    if (
        not isinstance(decimal_places, dict)
        or not decimal_places
        or any(not isinstance(g, str) or not g.strip() for g in decimal_places)
        or not isinstance(left_label, str)
        or not left_label.strip()
        or not isinstance(right_label, str)
        or not right_label.strip()
        or left_label == right_label
    ):
        raise SheetOpsError(
            "invalid_field", "Declare groups and two distinct side labels", {}
        )
    table, source, rejection = _read(
        path, "source", [group, side], amount, encoding, delimiter
    )
    sources = {"source": source}
    if rejection is not None:
        return _finish(sources, [rejection], None, {"totalsSkipped": True})
    buckets = {g: {left_label: [], right_label: []} for g in decimal_places}
    for index, row in enumerate(table["rows"], 1):
        if row[group] not in buckets or row[side] not in {left_label, right_label}:
            raise SheetOpsError(
                "unassigned_record",
                "Record is outside declared groups or sides",
                {"record": index, "group": row[group], "side": row[side]},
            )
        buckets[row[group]][row[side]].append(row[amount])
    _require(
        sum(len(v) for sides in buckets.values() for v in sides.values())
        == source["rowCount"],
        "groupCoverage",
    )
    groups = []
    for group_name, places in sorted(decimal_places.items()):
        left, right = buckets[group_name][left_label], buckets[group_name][right_label]
        report = check_totals(left, right, places, rounding, source_tolerance)
        _verify_totals(left, right, report, places, rounding, source_tolerance)
        groups.append({"group": group_name, **report})
    return _finish(
        sources,
        [],
        {"groups": groups},
        {
            "independentTotals": True,
            "completeGroupAssignments": True,
        },
    )


def _verify_totals(
    left: list[str],
    right: list[str],
    report: dict,
    places: int,
    rounding: str,
    tolerance: str,
) -> None:
    _require(
        type(report["leftCount"]) is int
        and type(report["rightCount"]) is int
        and (report["leftCount"], report["rightCount"]) == (len(left), len(right)),
        "totalsCounts",
    )
    _require(
        type(report["decimalPlaces"]) is int
        and report["decimalPlaces"] == places
        and report["rounding"] == rounding
        and report["sourceTolerance"] == tolerance,
        "totalsPolicy",
    )
    with localcontext(Context(prec=400)):
        quantum = Decimal(1).scaleb(-places)
        mode = {"half_up": ROUND_HALF_UP, "half_even": ROUND_HALF_EVEN}[rounding]
        for phase, round_lines, allowed in [
            ("source", False, Decimal(tolerance)),
            ("output", True, Decimal(0)),
        ]:
            totals = []
            for population in (left, right):
                values = [Decimal(value) for value in population]
                if round_lines:
                    values = [
                        value.quantize(quantum, rounding=mode) for value in values
                    ]
                totals.append(sum(values, Decimal(0)))
            difference = totals[0] - totals[1]
            actual = report[phase]
            _require(Decimal(actual["leftTotal"]) == totals[0], phase + ".leftTotal")
            _require(Decimal(actual["rightTotal"]) == totals[1], phase + ".rightTotal")
            _require(Decimal(actual["difference"]) == difference, phase + ".difference")
            _require(
                actual["matches"] is (abs(difference) <= allowed), phase + ".matches"
            )
        _require(
            report["validated"]
            is (report["source"]["matches"] and report["output"]["matches"]),
            "totalsValidated",
        )
