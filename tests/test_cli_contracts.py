import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook

from zalos_sheetops.cli import app
from zalos_sheetops.testing import CliRunner

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _parse_output(output: str) -> dict:
    return json.loads(output)


def _write_inspection_fixture(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet["A1"] = "Amount"
    sheet["B1"] = "Formula"
    sheet["A2"] = 123
    sheet["B2"] = "=A2*2"

    hidden = workbook.create_sheet("Hidden")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "Internal"

    workbook.save(path)


def _write_unresolved_drawing_relationship_fixture(path: Path) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        archive.writestr(
            "xl/workbook.xml",
            f"""
            <workbook xmlns="{MAIN_NS}" xmlns:r="{RELATIONSHIP_NS}">
              <sheets>
                <sheet name="Data" sheetId="1" r:id="rIdSheet1"/>
              </sheets>
            </workbook>
            """,
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            f"""
            <Relationships xmlns="{PACKAGE_RELATIONSHIP_NS}">
              <Relationship Id="rIdSheet1"
                Type="{RELATIONSHIP_NS}/worksheet"
                Target="worksheets/sheet1.xml"/>
            </Relationships>
            """,
        )
        archive.writestr(
            "xl/worksheets/sheet1.xml",
            f"""
            <worksheet xmlns="{MAIN_NS}" xmlns:r="{RELATIONSHIP_NS}">
              <dimension ref="A1"/>
              <sheetData/>
              <drawing r:id="missingDrawing"/>
            </worksheet>
            """,
        )


def test_schema_returns_machine_readable_command_manifest() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    assert payload["ok"] is True
    assert payload["command"] == "schema"
    assert payload["warnings"] == []
    assert payload["artifacts"] == []
    assert {command["name"] for command in payload["data"]["commands"]} == {
        "schema",
        "inspect",
        "read-table",
        "cells read",
        "cells search",
        "hyperlinks assert",
        "styles assert",
        "styles read",
        "validate",
        "formulas scan",
        "display scan",
        "compare structure",
        "compare labeled-cells",
        "normalize preview",
        "doctor libreoffice",
        "recalc",
        "diff",
        "reconcile",
    }


def test_schema_includes_json_contract_for_every_command() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    commands = {command["name"]: command for command in payload["data"]["commands"]}
    assert commands["inspect"]["input"] == {
        "path": "string",
        "engine": "auto | xml | openpyxl",
        "maxFormulaCells": "positive integer | null",
        "maxMergedRanges": "positive integer | null",
        "maxGraphics": "positive integer | null",
        "maxDefinedNames": "positive integer | null",
    }
    assert commands["read-table"]["input"] == {
        "path": "string",
        "sheet": "string | null",
        "headerStrategy": "detect | row",
        "headerRow": "integer | null",
        "maxRows": "positive integer | null",
    }
    assert commands["cells read"]["input"] == {
        "path": "string",
        "sheet": "string",
        "range": "A1-style range",
        "include": "list[values | formulas | cachedValues | numberFormats | dataTypes]",
        "engine": "auto | xml | openpyxl",
    }
    assert commands["cells search"]["input"] == {
        "path": "string",
        "query": "non-empty string",
        "sheet": "string | null",
        "range": "A1-style range | null",
        "fields": (
            "non-empty list[values | formulas | cachedValues | "
            "numberFormats | fontColors | fillColors]"
        ),
        "regex": "boolean",
        "caseSensitive": "boolean",
        "includeHidden": "boolean",
        "limit": "positive integer | null",
        "offset": "non-negative integer | null",
        "maxCells": "positive integer | null",
    }
    assert commands["hyperlinks assert"]["input"] == {
        "path": "string",
        "sheet": "string",
        "checks": "non-empty list[hyperlinkCheck]",
    }
    assert commands["styles assert"]["input"] == {
        "path": "string",
        "sheet": "string",
        "checks": "non-empty list[styleCheck]",
    }
    assert commands["styles read"]["input"] == {
        "path": "string",
        "sheet": "string",
        "range": "A1-style range",
        "maxCells": "positive integer | null",
    }
    assert commands["validate"]["input"] == {
        "path": "string",
        "requiredSheets": "list[string]",
        "requiredColumns": "dict[string, list[string]]",
        "uniqueKeys": "dict[string, list[string]]",
        "sheets": "list[sheetValidationRule] | null",
    }
    assert commands["formulas scan"]["input"] == {
        "path": "string",
        "sheet": "string | null",
        "range": "A1-style range | null",
        "maxResults": "positive integer | null",
        "engine": "auto | xml | openpyxl",
        "include": "list[formulas | errors] | null",
        "failFast": "boolean",
    }
    assert commands["display scan"]["input"] == {
        "path": "string",
        "sheet": "string | null",
        "range": "A1-style range | null",
        "maxResults": "positive integer | null",
        "maxCells": "positive integer | null",
        "includeHidden": "boolean",
    }
    assert commands["compare structure"]["input"] == {
        "expected": "workbookSource",
        "actual": "workbookSource",
        "compareFormulaCounts": "boolean",
    }
    assert commands["compare labeled-cells"]["input"] == {
        "expected": "workbookSheetSource",
        "actual": "workbookSheetSource",
        "labelColumn": "column letter",
        "valueColumn": "column letter",
        "labels": "non-empty list[string]",
        "compareFormulas": "boolean",
        "amountTolerance": "decimal string",
    }
    assert commands["doctor libreoffice"]["input"] == {}


def test_schema_marks_validate_as_available() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    commands = {command["name"]: command for command in payload["data"]["commands"]}
    assert commands["validate"]["status"] == "available"


def test_schema_marks_formulas_scan_as_available() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    commands = {command["name"]: command for command in payload["data"]["commands"]}
    assert commands["formulas scan"]["status"] == "available"


def test_schema_marks_display_scan_as_available() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    commands = {command["name"]: command for command in payload["data"]["commands"]}
    assert commands["display scan"]["status"] == "available"


def test_schema_includes_next_milestone_commands() -> None:
    runner = CliRunner()

    result = runner.invoke(app, ["schema", "--raw-output"])

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    commands = {command["name"]: command for command in payload["data"]["commands"]}
    assert commands["normalize preview"]["status"] == "available"
    assert commands["doctor libreoffice"]["status"] == "available"
    assert commands["recalc"]["status"] == "available"
    assert commands["recalc"]["input"] == {
        "path": "string",
        "outputDir": "string",
        "filename": "string | null",
        "executable": "string | null",
    }
    assert commands["diff"]["status"] == "available"
    assert commands["reconcile"]["status"] == "available"
    assert commands["normalize preview"]["input"]["columns"] == "list[columnRule]"
    assert commands["diff"]["input"]["keyColumns"] == "list[string]"
    assert commands["reconcile"]["input"]["amountTolerance"] == "decimal string"


def test_inspect_returns_compact_workbook_metadata(tmp_path: Path) -> None:
    workbook_path = tmp_path / "sample.xlsx"
    _write_inspection_fixture(workbook_path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "inspect",
            "--json",
            json.dumps({"path": str(workbook_path)}),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    assert payload["ok"] is True
    assert payload["command"] == "inspect"
    assert payload["warnings"] == []
    assert payload["data"]["workbook"]["engine"] == "xml"
    assert payload["data"]["workbook"]["sheetCount"] == 2
    assert payload["data"]["workbook"]["formulaCount"] == 1
    assert payload["data"]["sheets"] == [
        {
            "name": "Data",
            "maxRow": 2,
            "maxColumn": 2,
            "hidden": False,
            "state": "visible",
            "formulaCount": 1,
            "formulaCells": ["B2"],
            "formulaCellsTruncated": False,
            "dimensions": "A1:B2",
            "mergedRangeCount": 0,
            "mergedRanges": [],
            "mergedRangesTruncated": False,
            "hyperlinkCount": 0,
            "graphicCount": 0,
            "graphics": [],
            "graphicsTruncated": False,
        },
        {
            "name": "Hidden",
            "maxRow": 1,
            "maxColumn": 1,
            "hidden": True,
            "state": "hidden",
            "formulaCount": 0,
            "formulaCells": [],
            "formulaCellsTruncated": False,
            "dimensions": "A1:A1",
            "mergedRangeCount": 0,
            "mergedRanges": [],
            "mergedRangesTruncated": False,
            "hyperlinkCount": 0,
            "graphicCount": 0,
            "graphics": [],
            "graphicsTruncated": False,
        },
    ]


def test_inspect_accepts_xml_engine_and_formula_cell_limit(tmp_path: Path) -> None:
    workbook_path = tmp_path / "sample.xlsx"
    _write_inspection_fixture(workbook_path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "inspect",
            "--json",
            json.dumps(
                {
                    "path": str(workbook_path),
                    "engine": "xml",
                    "maxFormulaCells": 1,
                }
            ),
            "--raw-output",
        ],
    )

    assert result.exit_code == 0
    payload = _parse_output(result.output)
    assert payload["ok"] is True
    assert payload["data"]["workbook"]["engine"] == "xml"
    assert payload["data"]["workbook"]["formulaCount"] == 1


def test_inspect_missing_file_returns_stable_json_error(tmp_path: Path) -> None:
    missing_path = tmp_path / "missing.xlsx"
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "inspect",
            "--json",
            json.dumps({"path": str(missing_path)}),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    payload = _parse_output(result.output)
    assert payload == {
        "ok": False,
        "command": "inspect",
        "errorCode": "file_not_found",
        "message": "Workbook file does not exist",
        "details": {"path": str(missing_path)},
        "warnings": [],
        "artifacts": [],
    }


def test_inspect_invalid_workbook_returns_stable_json_error(tmp_path: Path) -> None:
    broken_path = tmp_path / "broken.xlsx"
    broken_path.write_text("not really an xlsx", encoding="utf-8")
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "inspect",
            "--json",
            json.dumps({"path": str(broken_path)}),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    payload = _parse_output(result.output)
    assert payload == {
        "ok": False,
        "command": "inspect",
        "errorCode": "invalid_workbook",
        "message": "Workbook could not be opened",
        "details": {"path": str(broken_path)},
        "warnings": [],
        "artifacts": [],
    }


def test_inspect_malformed_relationship_returns_stable_json_error(
    tmp_path: Path,
) -> None:
    broken_path = tmp_path / "broken-relationship.xlsx"
    _write_unresolved_drawing_relationship_fixture(broken_path)
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "inspect",
            "--json",
            json.dumps({"path": str(broken_path)}),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    payload = _parse_output(result.output)
    assert payload == {
        "ok": False,
        "command": "inspect",
        "errorCode": "invalid_workbook",
        "message": "Workbook could not be opened",
        "details": {"path": str(broken_path)},
        "warnings": [],
        "artifacts": [],
    }


def test_inspect_rejects_agent_hostile_paths() -> None:
    runner = CliRunner()

    result = runner.invoke(
        app,
        [
            "inspect",
            "--json",
            json.dumps({"path": "../secrets.xlsx"}),
            "--raw-output",
        ],
    )

    assert result.exit_code == 1
    payload = _parse_output(result.output)
    assert payload["ok"] is False
    assert payload["command"] == "inspect"
    assert payload["errorCode"] == "invalid_path"
    assert payload["message"] == "Path contains disallowed characters or traversal"
