from pathlib import Path

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.xlsx_xml import RowData, XlsxWorkbook
from zalos_sheetops.normalize.text import canonical_header


def read_table(
    path: Path,
    sheet_name: str | None,
    header_strategy: str,
    header_row: int | None,
    max_rows: int | None = None,
    value_mode: str = "raw",
) -> dict[str, object]:
    from zalos_sheetops.io.typed_values import validate_value_mode

    validate_value_mode(value_mode)
    with XlsxWorkbook(path, value_mode=value_mode) as workbook:
        resolved_sheet = _resolve_sheet(workbook, sheet_name)
        scan = _scan_table_rows(
            workbook,
            resolved_sheet,
            header_strategy,
            header_row,
            max_rows,
        )
    headers = _normalize_headers(scan["headerRow"].values)
    width = len(headers)
    rows: list[list[str]] = []
    for row in scan["rows"]:
        values = ["" if value is None else str(value) for value in row[:width]]
        if len(values) < width:
            values.extend([""] * (width - len(values)))
        rows.append(values)

    return {
        "sheet": resolved_sheet,
        "headerRow": scan["headerRow"].number,
        "headers": headers,
        "rows": rows,
        "rowCount": scan["rowCount"],
        "returnedRowCount": len(rows),
        "truncated": scan["truncated"],
    }


def _is_blank_value(value: object) -> bool:
    if value is None:
        return True
    if type(value) is str:
        return not value.strip()
    return False


def _scan_table_rows(
    workbook: XlsxWorkbook,
    sheet_name: str,
    header_strategy: str,
    header_row: int | None,
    max_rows: int | None,
) -> dict[str, object]:
    if header_strategy not in {"detect", "row"}:
        raise SheetOpsError(
            "invalid_header_strategy",
            "Header strategy must be 'detect' or 'row'",
            {"headerStrategy": header_strategy},
        )
    if header_strategy == "row":
        return _scan_explicit_header_table(workbook, sheet_name, header_row, max_rows)
    return _scan_detect_header_table(workbook, sheet_name, max_rows)


def _scan_explicit_header_table(
    workbook: XlsxWorkbook,
    sheet_name: str,
    header_row: int | None,
    max_rows: int | None,
) -> dict[str, object]:
    if header_row is None or header_row < 1:
        max_row = _sheet_max_row(workbook, sheet_name)
        raise SheetOpsError(
            "invalid_header_row",
            "Header row is outside the sheet bounds",
            {"headerRow": header_row, "sheet": sheet_name, "maxRow": max_row},
        )
    header_data: RowData | None = None
    rows: list[list[object]] = []
    row_count = 0
    max_row = 0
    stopped_early = False
    for row in workbook.iter_rows(sheet_name):
        if row.number > max_row:
            max_row = row.number
        if row.number == header_row:
            header_data = row
            continue
        if row.number <= header_row:
            continue
        if all(_is_blank_value(value) for value in row.values):
            continue
        if max_rows is not None and len(rows) >= max_rows:
            if header_data is not None:
                # The next data row proves truncation; the exact total is
                # unknown because the scan stops here.
                stopped_early = True
                break
            row_count += 1
            continue
        row_count += 1
        if header_data is not None:
            _check_row_width(row, header_data, sheet_name)
        rows.append(row.values)

    if header_data is None:
        raise SheetOpsError(
            "invalid_header_row",
            "Header row is outside the sheet bounds",
            {"headerRow": header_row, "sheet": sheet_name, "maxRow": max_row},
        )
    return {
        "headerRow": header_data,
        "rows": rows,
        "rowCount": None if stopped_early else row_count,
        "truncated": stopped_early or row_count > len(rows),
    }


def _scan_detect_header_table(
    workbook: XlsxWorkbook,
    sheet_name: str,
    max_rows: int | None,
) -> dict[str, object]:
    candidate_rows: list[int] = []
    best_row: RowData | None = None
    best_score = -1
    selected_row: RowData | None = None
    rows: list[list[object]] = []
    row_count = 0

    for row in workbook.iter_rows(sheet_name):
        nonblank = _is_nonblank_row(row.values)
        plausible = nonblank and _is_plausible_detect_header_row(row.values)
        if plausible:
            candidate_rows.append(row.number)
            if selected_row is None:
                selected_row = row

        if selected_row is not None and row.number > selected_row.number:
            if nonblank:
                row_count += 1
                if max_rows is None or len(rows) < max_rows:
                    _check_row_width(row, selected_row, sheet_name)
                    rows.append(row.values)
            continue

        if not nonblank:
            continue

        headers = [canonical_header(value) for value in row.values]
        non_blank = {header for header in headers if header}
        score = len(non_blank)
        if score > best_score:
            best_score = score
            best_row = row

    if len(candidate_rows) > 1:
        raise SheetOpsError(
            "ambiguous_header_row",
            "Multiple header rows are plausible in detect mode",
            {
                "sheet": sheet_name,
                "headerStrategy": "detect",
                "headerRows": candidate_rows,
            },
        )
    if selected_row is not None:
        return {
            "headerRow": selected_row,
            "rows": rows,
            "rowCount": row_count,
            "truncated": row_count > len(rows),
        }
    if best_row is None:
        raise SheetOpsError(
            "header_not_found",
            "Unable to detect header row",
            {"sheet": sheet_name, "headerStrategy": "detect"},
        )
    return _scan_explicit_header_table(workbook, sheet_name, best_row.number, max_rows)


def _resolve_sheet(
    workbook: XlsxWorkbook,
    sheet_name: str | None,
) -> str:
    if sheet_name is None:
        return workbook.worksheets[0].name
    workbook.worksheet(sheet_name)
    return sheet_name


def _check_row_width(row: RowData, header: RowData, sheet_name: str) -> None:
    if any(not _is_blank_value(value) for value in row.values[len(header.values) :]):
        raise SheetOpsError(
            "table_width_mismatch",
            "Data extends beyond the declared headers",
            {
                "sheet": sheet_name,
                "row": row.number,
                "headerColumns": len(header.values),
            },
        )


def _sheet_max_row(workbook: XlsxWorkbook, sheet_name: str) -> int:
    max_row = 0
    for row in workbook.iter_rows(sheet_name):
        if row.number > max_row:
            max_row = row.number
    return max_row


def _normalize_headers(header_row: list[object]) -> list[str]:
    headers = [canonical_header(value) for value in header_row]
    duplicates = _duplicate_headers(headers)
    blank_columns = [index + 1 for index, header in enumerate(headers) if not header]
    if duplicates or blank_columns:
        raise SheetOpsError(
            "invalid_headers",
            "Header row contains duplicate or blank columns",
            {"duplicates": duplicates, "blankColumns": blank_columns},
        )
    return headers


def _duplicate_headers(headers: list[str]) -> list[str]:
    seen: set[str] = set()
    duplicates = []
    for header in headers:
        if not header:
            continue
        if header in seen and header not in duplicates:
            duplicates.append(header)
            continue
        seen.add(header)
    return duplicates


def _is_plausible_detect_header_row(row: list[object]) -> bool:
    non_blank = [header for value in row if (header := canonical_header(value))]
    if len(non_blank) < 2:
        return False
    if len(set(non_blank)) != len(non_blank):
        return False
    return all(any(map(str.isalpha, header)) for header in non_blank)


def _is_nonblank_row(row: list[object]) -> bool:
    return any(not _is_blank_value(value) for value in row)
