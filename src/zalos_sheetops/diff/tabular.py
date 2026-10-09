from decimal import Decimal, InvalidOperation

from zalos_sheetops.errors import SheetOpsError


def diff_tables(
    left: dict[str, object],
    right: dict[str, object],
    key_columns: list[str],
    compare_columns: list[str] | None,
    amount_tolerance: str,
) -> dict[str, object]:
    _require_non_empty_columns(
        key_columns,
        "key_columns",
        "non-empty list[string]",
    )
    if compare_columns is not None:
        _require_non_empty_columns(
            compare_columns,
            "compare_columns",
            "non-empty list[string] | null",
        )
    left_rows = _rows(left)
    right_rows = _rows(right)
    left_by_key = _index_by_key(left_rows, key_columns, "left")
    right_by_key = _index_by_key(right_rows, key_columns, "right")
    compare = compare_columns
    if compare is None:
        compare = _all_compare_columns(
            left_rows,
            right_rows,
            key_columns,
        )
    else:
        _require_compare_columns(left_rows, compare, "left")
        _require_compare_columns(right_rows, compare, "right")

    left_keys = set(left_by_key)
    right_keys = set(right_by_key)
    added = [
        {"key": _key_dict(key_columns, key), "right": right_by_key[key]}
        for key in sorted(right_keys - left_keys)
    ]
    removed = [
        {"key": _key_dict(key_columns, key), "left": left_by_key[key]}
        for key in sorted(left_keys - right_keys)
    ]
    changed = _changed_rows(
        left_by_key,
        right_by_key,
        sorted(left_keys & right_keys),
        key_columns,
        compare,
        amount_tolerance,
    )

    return {
        "same": not added and not removed and not changed,
        "added": added,
        "removed": removed,
        "changed": changed,
        "summary": {
            "added": len(added),
            "removed": len(removed),
            "changed": len(changed),
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
    if not isinstance(rows, list):
        return []
    return [row for row in rows if isinstance(row, dict)]


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


def _require_compare_columns(
    rows: list[dict[str, object]],
    compare_columns: list[str],
    side: str,
) -> None:
    missing = []
    for row in rows:
        for column in compare_columns:
            if column not in row and column not in missing:
                missing.append(column)
    if missing:
        raise SheetOpsError(
            "missing_compare_column",
            "Required compare column is missing from a row",
            {"side": side, "columns": missing},
        )


def _changed_rows(
    left_by_key: dict[tuple[str, ...], dict[str, object]],
    right_by_key: dict[tuple[str, ...], dict[str, object]],
    shared_keys: list[tuple[str, ...]],
    key_columns: list[str],
    compare_columns: list[str],
    amount_tolerance: str,
) -> list[dict[str, object]]:
    changed = []
    for key in shared_keys:
        changes = []
        left_row = left_by_key[key]
        right_row = right_by_key[key]
        for column in compare_columns:
            left_value = left_row.get(column, "")
            right_value = right_row.get(column, "")
            if not _same_value(left_value, right_value, amount_tolerance):
                changes.append(
                    {"column": column, "left": left_value, "right": right_value}
                )
        if changes:
            changed.append({"key": _key_dict(key_columns, key), "changes": changes})
    return changed


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


def _key_dict(key_columns: list[str], key: tuple[str, ...]) -> dict[str, str]:
    return {column: key[index] for index, column in enumerate(key_columns)}


def _all_compare_columns(
    left_rows: list[dict[str, object]],
    right_rows: list[dict[str, object]],
    key_columns: list[str],
) -> list[str]:
    columns = []
    for row in [*left_rows, *right_rows]:
        for column in row:
            if column not in key_columns and column not in columns:
                columns.append(column)
    return columns


def _same_value(left: object, right: object, amount_tolerance: str) -> bool:
    if str(left) == str(right):
        return True

    left_decimal = _to_decimal(left)
    right_decimal = _to_decimal(right)
    if left_decimal is None or right_decimal is None:
        return False

    return abs(left_decimal - right_decimal) <= Decimal(amount_tolerance)


def _to_decimal(value: object) -> Decimal | None:
    try:
        decimal = Decimal(str(value))
    except InvalidOperation:
        return None
    if not decimal.is_finite():
        return None
    return decimal
