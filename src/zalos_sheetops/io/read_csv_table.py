import csv
from io import StringIO
from pathlib import Path

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_table import _normalize_headers


def read_csv_table(path: Path) -> dict[str, object]:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as error:
        raise SheetOpsError(
            "invalid_csv_encoding",
            "CSV is not valid UTF-8",
            {"path": str(path)},
        ) from error

    rows = list(csv.reader(StringIO(text)))
    if not rows or _rows_are_empty(rows):
        raise SheetOpsError("empty_csv", "CSV file has no rows", {"path": str(path)})

    headers = _normalize_headers(rows[0])
    records = []
    for row in rows[1:]:
        if all(value.strip() == "" for value in row):
            continue
        records.append(
            {
                header: row[index] if index < len(row) else ""
                for index, header in enumerate(headers)
            }
        )

    return {
        "sheet": "",
        "headerRow": 1,
        "headers": headers,
        "rows": records,
        "rowCount": len(records),
    }


def _rows_are_empty(rows: list[list[str]]) -> bool:
    return all(all(value.strip() == "" for value in row) for row in rows)
