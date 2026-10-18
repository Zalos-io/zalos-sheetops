from decimal import Context, Decimal, InvalidOperation, localcontext
from math import isfinite

from zalos_sheetops.errors import SheetOpsError


def reconcile_exact(
    left: dict[str, object],
    right: dict[str, object],
    left_key_columns: list[str],
    right_key_columns: list[str],
    left_amount_column: str,
    right_amount_column: str,
    amount_tolerance: str,
) -> dict[str, object]:
    _require_non_empty_columns(
        left_key_columns,
        "left_key_columns",
        "non-empty list[string]",
    )
    _require_non_empty_columns(
        right_key_columns,
        "right_key_columns",
        "non-empty list[string]",
    )
    if len(left_key_columns) != len(right_key_columns):
        raise SheetOpsError(
            "invalid_field",
            "Key column lists must have the same length",
            {"field": "key_columns", "expected": "same length"},
        )

    tolerance = _to_decimal(amount_tolerance, "amount_tolerance")
    if tolerance < 0:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "amount_tolerance", "expected": "decimal string"},
        )

    left_rows = _rows(left, "left")
    right_rows = _rows(right, "right")
    left_by_key = _index_by_key(
        left_rows, left_key_columns, "left", left_amount_column, "leftAmountColumn"
    )
    right_by_key = _index_by_key(
        right_rows, right_key_columns, "right", right_amount_column, "rightAmountColumn"
    )

    left_keys = set(left_by_key)
    right_keys = set(right_by_key)
    matched = []
    variance = []

    # Match the bounded decimal contract without inheriting caller precision/traps.
    with localcontext(Context(prec=400)):
        for key in sorted(left_keys & right_keys):
            left_row, left_amount = left_by_key[key]
            right_row, right_amount = right_by_key[key]
            difference = left_amount - right_amount
            key_data = _key_dict(left_key_columns, key)
            if abs(difference) <= tolerance:
                matched.append({"key": key_data, "left": left_row, "right": right_row})
            else:
                variance.append(
                    {
                        "key": key_data,
                        "leftAmount": str(left_amount),
                        "rightAmount": str(right_amount),
                        "difference": str(difference),
                    }
                )

    left_unmatched = [
        {"key": _key_dict(left_key_columns, key), "left": left_by_key[key][0]}
        for key in sorted(left_keys - right_keys)
    ]
    right_unmatched = [
        {"key": _key_dict(right_key_columns, key), "right": right_by_key[key][0]}
        for key in sorted(right_keys - left_keys)
    ]

    return {
        "matched": matched,
        "leftUnmatched": left_unmatched,
        "rightUnmatched": right_unmatched,
        "variance": variance,
        "summary": {
            "matched": len(matched),
            "leftUnmatched": len(left_unmatched),
            "rightUnmatched": len(right_unmatched),
            "variance": len(variance),
        },
    }


def _require_non_empty_columns(
    columns: list[str],
    field_name: str,
    expected: str,
) -> None:
    if columns:
        return
    raise SheetOpsError(
        "invalid_field",
        "Field must not be empty",
        {"field": field_name, "expected": expected},
    )


def _rows(table: dict[str, object], side: str) -> list[dict[str, object]]:
    rows = table.get("rows", [])
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise SheetOpsError(
            "invalid_field",
            "Table rows must be a list of objects",
            {"field": "rows", "expected": "list[object]"},
        )
    if table.get("truncated") or (
        "rowCount" in table and table["rowCount"] != len(rows)
    ):
        raise SheetOpsError(
            "incomplete_population",
            "Reconciliation requires the complete source population",
            {"side": side},
        )
    return rows


def _index_by_key(
    rows: list[dict[str, object]],
    key_columns: list[str],
    side: str,
    amount_column: str,
    amount_field: str,
) -> dict[tuple[str, ...], tuple[dict[str, object], Decimal]]:
    indexed = {}
    for row_number, row in enumerate(rows, start=1):
        key = _key(row, key_columns, side, row_number)
        if key in indexed:
            raise SheetOpsError(
                "duplicate_key",
                "Composite key is not unique",
                {"side": side, "key": _key_dict(key_columns, key)},
            )
        indexed[key] = (row, _row_decimal(row, amount_column, amount_field))
    return indexed


def _key(
    row: dict[str, object],
    key_columns: list[str],
    side: str,
    row_number: int,
) -> tuple[str, ...]:
    missing = [column for column in key_columns if column not in row]
    if missing:
        raise SheetOpsError(
            "missing_key_column",
            "Required key column is missing from a row",
            {"side": side, "columns": missing},
        )
    for column in key_columns:
        value = row[column]
        if (
            value is None
            or (isinstance(value, str) and not value.strip())
            or (isinstance(value, float) and not isfinite(value))
            or (isinstance(value, Decimal) and not value.is_finite())
        ):
            raise SheetOpsError(
                "invalid_key_value",
                "Reference keys must not be blank or nonfinite",
                {"side": side, "column": column, "row": row_number},
            )
    return tuple(str(row[column]) for column in key_columns)


def _key_dict(columns: list[str], key: tuple[str, ...]) -> dict[str, str]:
    return {column: key[index] for index, column in enumerate(columns)}


def _row_decimal(
    row: dict[str, object],
    amount_column: str,
    field_name: str,
) -> Decimal:
    if amount_column not in row:
        raise SheetOpsError(
            "missing_amount_column",
            "Required amount column is missing from a row",
            {"field": field_name, "column": amount_column},
        )
    return _to_decimal(row[amount_column], field_name)


def _to_decimal(value: object, field_name: str) -> Decimal:
    try:
        text = str(value)
        if len(text) > 256:
            raise InvalidOperation
        decimal = Decimal(text, context=Context())
    except InvalidOperation as error:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "decimal string"},
        ) from error
    if not decimal.is_finite() or (
        abs(decimal.as_tuple().exponent) > 128 or len(decimal.as_tuple().digits) > 128
    ):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "decimal string"},
        )
    return decimal
