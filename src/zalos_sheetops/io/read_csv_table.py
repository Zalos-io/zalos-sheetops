"""CSV records with explicit decoding and bounded, lossless previews."""

import codecs
import csv
from pathlib import Path

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_table import _normalize_headers


def read_csv_table(
    path: Path,
    *,
    encoding: str = "utf-8-sig",
    delimiter: str = ",",
    max_rows: int | None = None,
    header_row: int = 1,
) -> dict[str, object]:
    encoding = _encoding(path, encoding)
    if not isinstance(delimiter, str) or len(delimiter) != 1 or delimiter in "\r\n\0":
        raise SheetOpsError(
            "invalid_field", "Delimiter must be one character", {"field": "delimiter"}
        )
    for name, value in (("maxRows", max_rows), ("headerRow", header_row)):
        if (value is None and name == "headerRow") or (
            value is not None and (type(value) is not int or value < 1)
        ):
            raise SheetOpsError(
                "invalid_field", "Expected a positive integer", {"field": name}
            )
    try:
        with path.open(encoding=encoding, newline="") as stream:
            reader = csv.reader(stream, delimiter=delimiter, strict=True)
            for _ in range(header_row - 1):
                next(reader, None)
            header = next(reader, None)
            _reject_null_bytes(header or [], path)
            if not header or all(not value.strip() for value in header):
                raise SheetOpsError(
                    "empty_csv", "CSV file has no header row", {"path": str(path)}
                )
            headers = _normalize_headers(header)
            records, source_rows = [], []
            truncated = False
            while True:
                source_row = reader.line_num + 1
                row = next(reader, None)
                if row is None:
                    break
                _reject_null_bytes(row, path)
                if all(not value.strip() for value in row):
                    continue
                if max_rows is not None and len(records) >= max_rows:
                    truncated = True
                    break
                if len(row) > len(headers):
                    raise SheetOpsError(
                        "invalid_csv_row",
                        "CSV row has extra fields",
                        {
                            "row": source_row,
                            "expected": len(headers),
                            "actual": len(row),
                        },
                    )
                records.append(
                    {
                        name: row[i] if i < len(row) else ""
                        for i, name in enumerate(headers)
                    }
                )
                source_rows.append(source_row)
    except LookupError as error:
        raise SheetOpsError(
            "invalid_field", "Encoding must be a text codec", {"field": "encoding"}
        ) from error
    except UnicodeError as error:
        raise SheetOpsError(
            "invalid_csv_encoding", "CSV decoding failed", {"path": str(path)}
        ) from error
    except csv.Error as error:
        raise SheetOpsError(
            "invalid_csv",
            "CSV parsing failed",
            {"path": str(path), "reason": str(error)},
        ) from error
    return {
        "sheet": "",
        "headerRow": header_row,
        "headers": headers,
        "rows": records,
        "rowCount": len(records),
        "sourceRows": source_rows,
        "encoding": encoding,
        "delimiter": delimiter,
        "truncated": truncated,
    }


def _encoding(path: Path, encoding: str) -> str:
    if not isinstance(encoding, str):
        raise SheetOpsError(
            "invalid_field", "Encoding must be a string", {"field": "encoding"}
        )
    if encoding == "auto":
        with path.open("rb") as stream:
            prefix = stream.read(4)
        if prefix.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
            return "utf-32"
        if prefix.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
            return "utf-16"
        return "utf-8-sig"
    try:
        codecs.lookup(encoding)
    except LookupError as error:
        raise SheetOpsError(
            "invalid_field", "Unknown encoding", {"field": "encoding"}
        ) from error
    return encoding


def _reject_null_bytes(row: list[str], path: Path) -> None:
    if any("\0" in value for value in row):
        raise SheetOpsError(
            "invalid_csv_encoding",
            "CSV contains null bytes; specify its encoding",
            {"path": str(path)},
        )
