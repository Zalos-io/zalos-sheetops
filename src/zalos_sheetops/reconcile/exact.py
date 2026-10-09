from decimal import Decimal, InvalidOperation

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

    left_rows = _rows(left)
    right_rows = _rows(right)
    left_by_key = _index_by_key(left_rows, left_key_columns, "left")
    right_by_key = _index_by_key(right_rows, right_key_columns, "right")

    left_keys = set(left_by_key)
    right_keys = set(right_by_key)
    matched = []
    variance = []

    for key in sorted(left_keys & right_keys):
        left_row = left_by_key[key]
        right_row = right_by_key[key]
        left_amount = _row_decimal(left_row, left_amount_column, "leftAmountColumn")
        right_amount = _row_decimal(right_row, right_amount_column, "rightAmountColumn")
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
        {"key": _key_dict(left_key_columns, key), "left": left_by_key[key]}
        for key in sorted(left_keys - right_keys)
    ]
    right_unmatched = [
        {"key": _key_dict(right_key_columns, key), "right": right_by_key[key]}
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


def _rows(table: dict[str, object]) -> list[dict[str, object]]:
    rows = table.get("rows", [])
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise SheetOpsError(
            "invalid_field",
            "Table rows must be a list of objects",
            {"field": "rows", "expected": "list[object]"},
        )
    return rows


def _index_by_key(
    rows: list[dict[str, object]],
    key_columns: list[str],
    side: str,
) -> dict[tuple[str, ...], dict[str, object]]:
    indexed = {}
    for row in rows:
        key = _key(row, key_columns, side)
        if key in indexed:
            raise SheetOpsError(
                "duplicate_key",
                "Composite key is not unique",
                {"side": side, "key": _key_dict(key_columns, key)},
            )
        indexed[key] = row
    return indexed


def _key(
    row: dict[str, object],
    key_columns: list[str],
    side: str,
) -> tuple[str, ...]:
    missing = [column for column in key_columns if column not in row]
    if missing:
        raise SheetOpsError(
            "missing_key_column",
            "Required key column is missing from a row",
            {"side": side, "columns": missing},
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
        decimal = Decimal(str(value))
    except InvalidOperation as error:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "decimal string"},
        ) from error
    if not decimal.is_finite():
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "decimal string"},
        )
    return decimal
