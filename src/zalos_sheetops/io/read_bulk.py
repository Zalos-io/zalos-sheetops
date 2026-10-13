"""Strict, typed bulk reads. Exact cells and formula work stay on the XML API."""

from __future__ import annotations

import re
from pathlib import Path

import fastexcel
import pyarrow as pa
import pyarrow.compute as pc

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_table import read_table
from zalos_sheetops.io.xlsx_xml import XlsxWorkbook

_FORMULA = re.compile(rb"<(?:[^<>\s/:]+:)?f(?=[\s/>])")


def read_bulk_table(
    path: Path,
    sheet_name: str,
    header_row: int,
    *,
    column_types: dict[str, str] | None = None,
) -> pa.RecordBatch:
    """Read a values-only XLSX table into Arrow using fastexcel.

    Sheet and one-based header row are explicit. Headers use SheetOps
    normalization. Interior blank rows remain in place; trailing empty rows
    outside the native reader's used range are omitted. Values are typed,
    unlike read_table's string-valued contract. Dates have millisecond
    precision. No formulas are evaluated or accepted, even with cached values.

    Infer from the entire column, rejecting mixed types and cell errors.
    column_types may explicitly convert named columns to strings; this does
    not recover formatting or leading zeros from numeric Excel cells.
    """
    if not isinstance(sheet_name, str) or not sheet_name:
        raise SheetOpsError("invalid_sheet", "An explicit sheet is required", {})
    if type(header_row) is not int or header_row < 1:
        raise SheetOpsError("invalid_header_row", "Use a positive header row", {})
    header = read_table(path, sheet_name, "row", header_row, max_rows=0)
    headers = header["headers"]
    overrides = _column_overrides(headers, column_types)
    _check_values_only(path, sheet_name)
    if overrides:
        _check_override_precision(path, sheet_name, header_row, overrides)
    try:
        reader = fastexcel.read_excel(path)
        sheet = reader.load_sheet(
            sheet_name,
            header_row=header_row - 1,
            skip_rows=0,
            schema_sample_rows=None,
            dtype_coercion="strict",
            dtypes=overrides or None,
        )
        batch, errors = sheet.to_arrow_with_errors()
    except fastexcel.UnsupportedColumnTypeCombinationError as exc:
        raise SheetOpsError(
            "bulk_type_conflict",
            "Mixed column types require an explicit string column override",
            {"sheet": sheet_name},
        ) from exc
    except fastexcel.FastExcelError as exc:
        raise SheetOpsError(
            "bulk_read_failed",
            "The bulk reader could not read the table",
            {"sheet": sheet_name, "reason": str(exc)[:500]},
        ) from exc
    if errors is not None:
        raise SheetOpsError(
            "bulk_cell_error",
            "Cell errors cannot be returned as empty values",
            {"sheet": sheet_name, "errors": str(errors)[:1000]},
        )
    if batch.num_columns != len(headers):
        raise SheetOpsError(
            "bulk_width_mismatch",
            "Data extends beyond the declared headers",
            {"headerColumns": len(headers), "dataColumns": batch.num_columns},
        )
    for column in batch.columns:
        if pa.types.is_integer(column.type) or pa.types.is_floating(column.type):
            unsafe = pc.or_(
                pc.greater_equal(column, 2**53), pc.less_equal(column, -(2**53))
            )
            if pc.any(unsafe).as_py():
                raise SheetOpsError(
                    "bulk_precision_risk",
                    "Large numeric values require exact XML reads; store IDs as text",
                    {"sheet": sheet_name},
                )
    return batch.rename_columns(headers)


def _check_override_precision(
    path: Path, sheet_name: str, header_row: int, overrides: dict[int, str]
) -> None:
    # Explicit coercion is uncommon and forfeits the fast path: verify the
    # original numeric values before a string cast can conceal rounding.
    with XlsxWorkbook(path) as workbook:
        for row in workbook.iter_rows(sheet_name):
            if row.number <= header_row:
                continue
            for index in overrides:
                value = row.values[index] if index < len(row.values) else None
                if type(value) in {int, float} and abs(value) >= 2**53:
                    raise SheetOpsError(
                        "bulk_precision_risk",
                        "A string override cannot recover numeric precision",
                        {"sheet": sheet_name, "row": row.number, "column": index + 1},
                    )


def _column_overrides(
    headers: list[str], column_types: dict[str, str] | None
) -> dict[int, str]:
    if column_types is None:
        return {}
    if not isinstance(column_types, dict) or any(
        name not in headers or dtype != "string" for name, dtype in column_types.items()
    ):
        raise SheetOpsError(
            "invalid_bulk_schema",
            "Overrides must map existing normalized headers to 'string'",
            {},
        )
    return {headers.index(name): dtype for name, dtype in column_types.items()}


def _check_values_only(path: Path, sheet_name: str) -> None:
    with XlsxWorkbook(path) as workbook:
        xml = workbook.archive.read(workbook.worksheet(sheet_name).path)
    # This conservative scan is not an XML parser. Native parsing still
    # validates the sheet; unsupported encodings/DTD cannot bypass the check.
    if b"\x00" in xml[:256] or xml.startswith((b"\xff\xfe", b"\xfe\xff")):
        raise SheetOpsError(
            "bulk_encoding_unsupported", "Use XML reads for this XML encoding", {}
        )
    if b"<!DOCTYPE" in xml or b"<!ENTITY" in xml:
        raise SheetOpsError(
            "bulk_encoding_unsupported", "Use XML reads for XML with declarations", {}
        )
    if _FORMULA.search(xml):
        raise SheetOpsError(
            "bulk_formulas_unsupported",
            "Formula sheets require XML reads and separate formula validation",
            {"sheet": sheet_name},
        )
