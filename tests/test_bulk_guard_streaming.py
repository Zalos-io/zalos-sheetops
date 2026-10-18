"""The preflight guard must not materialize the complete worksheet XML."""

import tracemalloc
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_bulk import _check_values_only


def worksheet_fixture(tmp_path: Path, xml: bytes) -> Path:
    path = tmp_path / "source.xlsx"
    workbook = Workbook()
    workbook.active.title = "Data"
    workbook.save(path)
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["xl/worksheets/sheet1.xml"] = xml
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    return path


def test_bulk_preflight_uses_bounded_memory_on_large_text_sheet(tmp_path: Path) -> None:
    xml = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>'
        + b"x" * (12 * 1024 * 1024)
        + b"</t></is></c></row></sheetData></worksheet>"
    )
    path = worksheet_fixture(tmp_path, xml)
    tracemalloc.start()
    try:
        _check_values_only(path, "Data")
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    assert peak < 8 * 1024 * 1024


@pytest.mark.parametrize("offset", [-3, -1, 2])
def test_bulk_preflight_rejects_formula_across_stream_boundaries(
    tmp_path: Path, offset: int
) -> None:
    start = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    )
    # Also exercise prefixed f tags rather than only openpyxl's usual spelling.
    formula = b'<n:f xmlns:n="http://schemas.openxmlformats.org/spreadsheetml/2006/main">1+1</n:f>'
    before = start + b"<!--"
    after = b'--><sheetData><row r="1"><c r="A1">'
    xml = (
        before
        + b"x" * (1024 * 1024 + offset - len(before) - len(after))
        + after
        + formula
        + b"</c></row></sheetData></worksheet>"
    )
    path = worksheet_fixture(tmp_path, xml)
    with pytest.raises(SheetOpsError) as error:
        _check_values_only(path, "Data")
    assert error.value.error_code == "bulk_formulas_unsupported"
