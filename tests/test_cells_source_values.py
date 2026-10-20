"""Original numeric lexemes accompany typed dates without another XML reader."""

import json
from pathlib import Path
from zipfile import ZipFile

import pytest

from zalos_sheetops.cli import app
from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_cells import read_cells
from zalos_sheetops.testing import CliRunner


def source_workbook(path: Path) -> None:
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    with ZipFile(path, "w") as archive:
        archive.writestr(
            "xl/workbook.xml",
            f'<workbook xmlns="{main}" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
        )
        archive.writestr(
            "xl/styles.xml",
            f'<styleSheet xmlns="{main}"><cellXfs count="2">'
            '<xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>',
        )
        archive.writestr(
            "xl/sharedStrings.xml",
            f'<sst xmlns="{main}"><si><t>000012345678901234567890</t></si></sst>',
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            f'<worksheet xmlns="{main}"><sheetData><row r="2">'
            '<c r="A2" t="s"><v>0</v></c>'
            '<c r="B2"><v>-123456789.012345678901234567890123</v></c>'
            '<c r="C2" s="1"><v>59.25</v></c>'
            '<c r="D2"><f>1/7</f><v>1.234567890123456789E-7</v></c>'
            '<c r="E2"><f>SUM(B2)</f></c>'
            '<c r="F2"><v>-0.0000</v></c>'
            "</row></sheetData></worksheet>",
        )


@pytest.mark.parametrize("value_mode", ["raw", "typed"])
def test_source_values_keep_original_digits_beside_typed_values(
    tmp_path: Path, value_mode: str
) -> None:
    path = tmp_path / "source.xlsx"
    source_workbook(path)
    before = path.read_bytes()
    result = read_cells(
        path,
        "Data",
        "A2:G2",
        ["values", "sourceValues", "formulas"],
        engine="xml",
        value_mode=value_mode,
    )
    cells = {cell["address"]: cell for cell in result["cells"]}
    assert cells["A2"]["sourceValue"] == "000012345678901234567890"
    assert cells["B2"]["sourceValue"] == "-123456789.012345678901234567890123"
    assert cells["C2"]["sourceValue"] == "59.25"
    assert cells["D2"]["sourceValue"] == "1.234567890123456789E-7"
    assert cells["D2"]["formula"] == "=1/7"
    assert cells["E2"]["sourceValue"] is None
    assert cells["E2"]["formula"] == "=SUM(B2)"
    assert cells["F2"]["sourceValue"] == "-0.0000"
    assert "G2" not in cells
    if value_mode == "typed":
        assert cells["C2"]["value"] == "1900-02-28T06:00:00"
        assert result["dateSystem"] == "1900"
    else:
        assert cells["C2"]["value"] == 59.25
        assert "dateSystem" not in result
    assert path.read_bytes() == before


def test_source_values_are_opt_in_and_require_xml(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    source_workbook(path)
    legacy = read_cells(path, "Data", "B2", engine="xml")
    assert "sourceValue" not in legacy["cells"][0]
    with pytest.raises(SheetOpsError) as caught:
        read_cells(path, "Data", "B2", ["sourceValues"], engine="openpyxl")
    assert caught.value.error_code == "invalid_field"
    assert caught.value.details["field"] == "engine"


def test_cells_cli_source_values_json_contract(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    source_workbook(path)
    result = CliRunner().invoke(
        app,
        [
            "cells",
            "read",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "sheet": "Data",
                    "range": "B2:C2",
                    "engine": "xml",
                    "valueMode": "typed",
                    "include": ["values", "sourceValues"],
                }
            ),
        ],
    )
    assert result.exit_code == 0
    receipt = json.loads(result.output)
    assert receipt["ok"] is True and receipt["command"] == "cells read"
    assert receipt["data"]["dateSystem"] == "1900"
    assert receipt["data"]["cells"] == [
        {
            "address": "B2",
            "value": -123456789.01234567,
            "sourceValue": "-123456789.012345678901234567890123",
        },
        {"address": "C2", "value": "1900-02-28T06:00:00", "sourceValue": "59.25"},
    ]
