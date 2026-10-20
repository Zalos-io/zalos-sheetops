"""Compact evidence from complete string tables; no file I/O or reconciliation claim."""

from decimal import Context, Decimal, localcontext

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.validation.totals import _decimal

_VALUE_PREVIEW_LIMIT = 80


def profile_table(
    table: dict[str, object],
    *,
    key_columns: list[str] | None = None,
    amount_columns: list[str] | None = None,
) -> dict[str, object]:
    """Validate a complete read_table/read_csv_table result and summarize it.

    Every row must match the declared headers. Keys must be nonblank text;
    amounts must be bounded finite decimal strings, retaining source precision.
    All invalid records are counted; at most 20 issue examples are returned.
    Issues include valueType and, for strings, an original valuePreview capped at
    80 characters with valueTruncated. Non-string values are never stringified.
    Column min/max are lexical; first/last retain source order. No uniqueness,
    cross-source balance, rounding, or fresh formula evaluation is implied.
    """
    headers, rows = _population(table)
    keys = _columns(key_columns, headers)
    amounts = _columns(amount_columns, headers)
    if not rows:
        raise SheetOpsError("empty_population", "A profile requires data rows", {})
    columns = {
        name: {"first": None, "last": None, "min": None, "max": None, "blankCount": 0}
        for name in headers
    }
    totals = {name: Decimal(0) for name in amounts}
    issues: list[dict[str, object]] = []
    issue_count = invalid_records = 0
    source_rows = table.get("sourceRows")
    with localcontext(Context(prec=400)):
        for index, row in enumerate(rows):
            values = _values(row, headers, index + 1)
            record_issues = _value_issues(values, keys, amounts, totals)
            if record_issues:
                invalid_records += 1
                issue_count += len(record_issues)
                for issue in record_issues:
                    if len(issues) < 20:
                        issues.append(
                            {
                                "record": index + 1,
                                **(
                                    {"sourceRow": source_rows[index]}
                                    if source_rows is not None
                                    else {}
                                ),
                                **issue,
                                **_value_evidence(values[issue["column"]]),
                            }
                        )
            for name, value in values.items():
                if not isinstance(value, str):
                    continue
                stats = columns[name]
                if index == 0:
                    stats["first"] = value
                stats["last"] = value
                if not value.strip():
                    stats["blankCount"] += 1
                else:
                    stats["min"] = (
                        value if stats["min"] is None else min(stats["min"], value)
                    )
                    stats["max"] = (
                        value if stats["max"] is None else max(stats["max"], value)
                    )
    if invalid_records:
        raise SheetOpsError(
            "invalid_table_values",
            "Complete population contains invalid values",
            {
                "invalidRecordCount": invalid_records,
                "issueCount": issue_count,
                "issues": issues,
                "issuesTruncated": issue_count > len(issues),
            },
        )
    return {
        "validated": True,
        "rowCount": len(rows),
        "columns": columns,
        "totals": {
            name: format(total.copy_abs() if total == 0 else total, "f")
            for name, total in totals.items()
        },
        "checks": {
            "completePopulation": True,
            "rectangularRows": True,
            "nonblankTextKeys": keys,
            "finiteDecimalAmounts": amounts,
        },
    }


def _population(table: dict[str, object]) -> tuple[list[str], list]:
    headers, rows = table.get("headers"), table.get("rows")
    if (
        not isinstance(headers, list)
        or not headers
        or any(not isinstance(name, str) or not name.strip() for name in headers)
        or len(set(headers)) != len(headers)
    ):
        raise SheetOpsError("invalid_headers", "Use unique nonblank string headers", {})
    if not isinstance(rows, list):
        raise SheetOpsError(
            "invalid_field", "Table rows must be a list", {"field": "rows"}
        )
    if (
        table.get("truncated") is not False
        or type(table.get("rowCount")) is not int
        or table["rowCount"] != len(rows)
        or type(table.get("returnedRowCount")) is not int
        or table["returnedRowCount"] != len(rows)
    ):
        raise SheetOpsError(
            "incomplete_population", "A profile requires complete row counts", {}
        )
    if "sourceRows" in table:
        positions = table["sourceRows"]
        if (
            not isinstance(positions, list)
            or len(positions) != len(rows)
            or any(type(value) is not int or value < 1 for value in positions)
            or any(left >= right for left, right in zip(positions, positions[1:]))
        ):
            raise SheetOpsError(
                "incomplete_population", "Invalid source-row receipt", {}
            )
    return headers, rows


def _columns(columns: list[str] | None, headers: list[str]) -> list[str]:
    if columns is None:
        return []
    if not isinstance(columns, list) or any(
        not isinstance(name, str) for name in columns
    ):
        raise SheetOpsError("invalid_field", "Columns must be a list of names", {})
    if len(set(columns)) != len(columns):
        raise SheetOpsError("invalid_field", "Validation columns must be unique", {})
    missing = sorted(set(columns) - set(headers))
    if missing:
        raise SheetOpsError(
            "missing_column", "Validation column is not declared", {"columns": missing}
        )
    return columns


def _values(row: object, headers: list[str], record: int) -> dict[str, object]:
    if isinstance(row, dict) and set(row) == set(headers):
        return row
    if isinstance(row, list) and len(row) == len(headers):
        return dict(zip(headers, row))
    raise SheetOpsError(
        "table_width_mismatch", "Row differs from declared headers", {"record": record}
    )


def _value_issues(
    values: dict[str, object],
    keys: list[str],
    amounts: list[str],
    totals: dict[str, Decimal],
) -> list[dict[str, object]]:
    issues = []
    for name, value in values.items():
        if not isinstance(value, str):
            issues.append({"column": name, "errorCode": "non_string_value"})
        elif name in keys and not value.strip():
            issues.append({"column": name, "errorCode": "invalid_key_value"})
        if name in amounts:
            try:
                totals[name] += _decimal(value, name)
            except SheetOpsError:
                issues.append({"column": name, "errorCode": "invalid_amount_value"})
    return issues


def _value_evidence(value: object) -> dict[str, object]:
    evidence: dict[str, object] = {"valueType": type(value).__name__}
    if isinstance(value, str):
        evidence.update(
            valuePreview=value[:_VALUE_PREVIEW_LIMIT],
            valueTruncated=len(value) > _VALUE_PREVIEW_LIMIT,
        )
    return evidence
