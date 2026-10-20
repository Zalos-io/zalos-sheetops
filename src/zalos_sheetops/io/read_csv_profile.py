"""Complete CSV validation with file identity and bounded original-source evidence."""

import hashlib
from pathlib import Path

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_csv_table import read_csv_table
from zalos_sheetops.validation.table_profile import profile_table


def read_csv_profile(
    path: Path,
    *,
    key_columns: list[str],
    amount_columns: list[str],
    encoding: str = "auto",
    delimiter: str = ",",
    header_row: int = 1,
    max_rows: int | None = None,
) -> tuple[dict[str, object], dict[str, object]]:
    """Return (complete_table, profile_receipt) with source path/hash/headers.

    max_rows must remain None; partial-source requests fail before file access.
    Explicit column lists select the same checks as profile_table; [] deliberately
    selects no key or amount checks. All records are read, never a row preview.
    Input rejections preserve their error_code/details and add source metadata.
    invalid_table_values also includes sourceExcerpts: original decoded CSV text
    at each retained issue's physical row, up to 3 lines / 240 characters. These
    excerpts read the original file separately from the validator. Complete bad
    record counts remain independent of the bounded examples. A changed source
    blocks the receipt. This does not prove uniqueness, balance or formula eval.
    """
    if max_rows is not None:
        raise SheetOpsError(
            "incomplete_population",
            "CSV validation requires the complete source",
            {"field": "maxRows"},
        )
    if not isinstance(key_columns, list) or not isinstance(amount_columns, list):
        raise SheetOpsError(
            "invalid_field",
            "Supply explicit validation column lists",
            {"field": "columns"},
        )
    path = Path(path)
    if not path.is_file():
        raise SheetOpsError(
            "file_not_found", "CSV source file does not exist", {"path": str(path)}
        )
    before = _source_digest(path)
    table = read_csv_table(
        path,
        encoding=encoding,
        delimiter=delimiter,
        header_row=header_row,
        max_rows=None,
    )
    source = {
        "path": str(path.resolve()),
        "sha256": before,
        "headers": table["headers"],
        "rowCount": table["rowCount"],
        "encoding": table["encoding"],
        "delimiter": table["delimiter"],
        "headerRow": table["headerRow"],
    }
    try:
        receipt = profile_table(
            table, key_columns=key_columns, amount_columns=amount_columns
        )
    except SheetOpsError as error:
        _require_same_source(path, before)
        details = {**error.details, "source": source}
        if error.error_code == "invalid_table_values":
            details["sourceExcerpts"] = _source_excerpts(
                path, table["encoding"], error.details["issues"]
            )
        _require_same_source(path, before)
        raise SheetOpsError(error.error_code, error.message, details) from error
    _require_same_source(path, before)
    return table, {**receipt, "source": source}


def _source_digest(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            while chunk := stream.read(1 << 20):
                digest.update(chunk)
    except OSError as error:
        raise SheetOpsError(
            "source_read_failed", "CSV source could not be read", {"path": str(path)}
        ) from error
    return digest.hexdigest()


def _require_same_source(path: Path, before: str) -> None:
    if _source_digest(path) != before:
        raise SheetOpsError(
            "source_changed",
            "CSV source changed during validation",
            {"path": str(path)},
        )


def _source_excerpts(
    path: Path, encoding: str, issues: list[dict]
) -> list[dict[str, object]]:
    starts = sorted({issue["sourceRow"] for issue in issues})
    wanted = {start + offset for start in starts for offset in range(3)}
    last_wanted = max(wanted)
    lines: dict[int, str] = {}
    oversized: set[int] = set()
    try:
        with path.open(encoding=encoding, newline="") as stream:
            for number, line in enumerate(stream, start=1):
                if number in wanted:
                    lines[number] = line[:241]
                    if len(line) > 240:
                        oversized.add(number)
                if number >= last_wanted:
                    break
    except (OSError, UnicodeError) as error:
        raise SheetOpsError(
            "source_read_failed",
            "CSV source evidence could not be read",
            {"path": str(path)},
        ) from error
    result = []
    for start in starts:
        numbers = [n for n in range(start, start + 3) if n in lines]
        if not numbers:
            raise SheetOpsError(
                "source_changed",
                "CSV source no longer contains the rejected row",
                {"path": str(path), "sourceRow": start},
            )
        text = "".join(lines[n] for n in numbers)
        result.append(
            {
                "sourceRow": start,
                "endSourceRow": max(numbers),
                "textPreview": text[:240],
                "textTruncated": len(text) > 240
                or bool(oversized.intersection(numbers)),
            }
        )
    return result
