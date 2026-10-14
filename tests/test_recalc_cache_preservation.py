"""Cache invalidation preserves literal OOXML content, including CDATA."""

from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import Workbook

from zalos_sheetops.delivery_cache import fresh_calculation_input
from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.recalc.fresh_input import without_formula_caches
from zalos_sheetops.recalc.libreoffice import recalc_workbook


@pytest.mark.parametrize("encoding", ["utf-16-le", "utf-16-be"])
def test_cache_removal_preserves_utf16_units(encoding: str) -> None:
    xml = (
        '<?xml version="1.0" encoding="UTF-16"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row><c r="A1"><f>1+1</f><v>999</v></c>'
        '<c r="B1"><f>1+1</f><v/></c></row></sheetData></worksheet>'
    )
    bom = b"\xff\xfe" if encoding.endswith("le") else b"\xfe\xff"
    expected = xml.replace("<v>999</v>", "").replace("<v/>", "")
    assert without_formula_caches(bom + xml.encode(encoding)) == bom + expected.encode(
        encoding
    )


def test_cache_removal_resolves_nonstandard_worksheet_target(tmp_path: Path) -> None:
    source, target = tmp_path / "source.xlsx", tmp_path / "target.xlsx"
    book = Workbook()
    book.active["A1"] = "=1+1"
    book.save(source)
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["xl/sheets/data.xml"] = parts.pop("xl/worksheets/sheet1.xml").replace(
        b"<v />", b"<v>999</v>"
    )
    parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace(
        b"/xl/worksheets/sheet1.xml", b"/xl/sheets/data.xml"
    )
    with ZipFile(source, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    fresh_calculation_input(source, target)
    with ZipFile(target) as archive:
        assert b"<v>999</v>" not in archive.read("xl/sheets/data.xml")


def test_literal_formula_markup_is_not_a_formula(tmp_path: Path) -> None:
    source, target = tmp_path / "source.xlsx", tmp_path / "target.xlsx"
    xml = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>'
        "<![CDATA[<f>not a formula</f>]]></t></is></c>"
        '<c r="B1"><f>1+1</f><v>999</v></c></row></sheetData></worksheet>'
    ).encode()
    Workbook().save(source)
    with ZipFile(source) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts["xl/worksheets/sheet1.xml"] = xml
    with ZipFile(source, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    fresh_calculation_input(source, target)
    with ZipFile(target) as archive:
        result = archive.read("xl/worksheets/sheet1.xml")
    assert (
        b'<c r="A1" t="inlineStr"><is><t><![CDATA[<f>not a formula</f>]]></t></is></c>'
        in result
    )
    assert b'<c r="B1"><f>1+1</f></c>' in result


def test_invalid_staged_worksheet_has_structured_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "corrupt.xlsx"
    with ZipFile(source, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", "<worksheet><")
    monkeypatch.setattr(
        "zalos_sheetops.recalc.libreoffice.detect_libreoffice",
        lambda _: {"available": True, "executable": "calc"},
    )
    with pytest.raises(SheetOpsError) as error:
        recalc_workbook(source, tmp_path / "output", None)
    assert error.value.error_code == "invalid_workbook"
