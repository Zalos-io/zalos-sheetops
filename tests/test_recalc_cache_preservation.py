"""Cache invalidation preserves literal OOXML content, including CDATA."""

from pathlib import Path
from zipfile import ZipFile

import pytest

from zalos_sheetops.delivery_cache import fresh_calculation_input
from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.recalc.libreoffice import recalc_workbook


def test_literal_formula_markup_is_not_a_formula(tmp_path: Path) -> None:
    source, target = tmp_path / "source.xlsx", tmp_path / "target.xlsx"
    xml = (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>'
        "<![CDATA[<f>not a formula</f>]]></t></is></c>"
        '<c r="B1"><f>1+1</f><v>999</v></c></row></sheetData></worksheet>'
    ).encode()
    with ZipFile(source, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", xml)
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
