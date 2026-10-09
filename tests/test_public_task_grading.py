from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.workbook.defined_name import DefinedName

from zalos_sheetops.benchmarks.public.grading import grade_workbook


def test_correct_answer_with_unrequested_edit_fails(tmp_path: Path) -> None:
    initial, golden, output = [
        tmp_path / name for name in ("input.xlsx", "golden.xlsx", "output.xlsx")
    ]
    book = Workbook()
    book.active["A1"] = "protected text"
    book.active["B1"] = 2
    book.save(initial)
    book.active["B1"] = 4
    book.save(golden)
    book.active["A1"] = "silently damaged"
    book.save(output)
    result = grade_workbook(initial, golden, output, {"answer_position": "B1"})
    assert result["status"] == "incorrect"
    assert result["preservationFailures"][0]["cell"] == "A1"


def test_missing_output_stays_a_failed_attempt(tmp_path: Path) -> None:
    result = grade_workbook(
        tmp_path / "input.xlsx", tmp_path / "golden.xlsx", tmp_path / "missing.xlsx", {}
    )
    assert result["status"] == "missing_output"


def test_unknown_formula_cache_is_not_scored_as_correct_blank(tmp_path: Path) -> None:
    paths = [tmp_path / name for name in ("input.xlsx", "golden.xlsx", "output.xlsx")]
    book = Workbook()
    book.active["A1"] = "=1+2"
    for path in paths:
        book.save(path)
    result = grade_workbook(*paths, {"answer_position": "A1"})
    assert result["status"] == "oracle_cache_missing"


def test_correct_public_answer_passes(tmp_path: Path) -> None:
    paths = [tmp_path / name for name in ("input.xlsx", "golden.xlsx", "output.xlsx")]
    book = Workbook()
    book.active["A1"] = 1
    book.save(paths[0])
    book.active["A1"] = 2
    book.save(paths[1])
    book.save(paths[2])
    assert grade_workbook(*paths, {"answer_position": "A1"})["status"] == "pass"
    loaded = load_workbook(paths[0])
    assert loaded.active["A1"].value == 1


def test_correct_values_without_required_structure_fail(tmp_path: Path) -> None:
    paths = [tmp_path / name for name in ("input.xlsx", "golden.xlsx", "output.xlsx")]
    book = Workbook()
    book.active["A1"] = 1
    book.save(paths[0])
    book.defined_names.add(DefinedName("Required", attr_text="'Sheet'!$A$1"))
    book.active.sheet_view.showGridLines = False
    book.active["A1"].fill = PatternFill("solid", fgColor="FFFF00")
    book.save(paths[1])
    book.defined_names.clear()
    book.active.sheet_view.showGridLines = True
    book.active["A1"].fill = PatternFill()
    book.save(paths[2])
    result = grade_workbook(
        *paths,
        {"answer_position": "A1", "instruction": "Fill A1 yellow and hide gridlines"},
    )
    assert result["status"] == "incorrect"
    assert {d["kind"] for d in result["structureDifferences"]} >= {
        "defined_names",
        "sheet_view",
    }
    assert any(d.get("kind") == "required_style" for d in result["valueDifferences"])


def test_corrupt_output_is_a_failed_attempt(tmp_path: Path) -> None:
    paths = [tmp_path / name for name in ("input.xlsx", "golden.xlsx", "output.xlsx")]
    book = Workbook()
    book.save(paths[0])
    book.save(paths[1])
    paths[2].write_text("not an XLSX")
    assert (
        grade_workbook(*paths, {"answer_position": "A1"})["status"] == "invalid_output"
    )


def test_evaluated_empty_string_is_not_a_missing_formula_cache(tmp_path: Path) -> None:
    from zipfile import ZipFile

    paths = [tmp_path / name for name in ("input.xlsx", "golden.xlsx", "output.xlsx")]
    for path in paths:
        book = Workbook()
        book.active["A1"] = '=IF(TRUE,"",1)'
        book.save(path)
        # LibreOffice marks an evaluated empty-string cache as t="str".
        with ZipFile(path) as archive:
            members = {name: archive.read(name) for name in archive.namelist()}
        member = "xl/worksheets/sheet1.xml"
        members[member] = members[member].replace(b'<c r="A1">', b'<c r="A1" t="str">')
        with ZipFile(path, "w") as archive:
            for name, data in members.items():
                archive.writestr(name, data)
    assert grade_workbook(*paths, {"answer_position": "A1"})["status"] == "pass"
    raw = tmp_path / "raw.xlsx"
    book = Workbook()
    book.active["A1"] = '=IF(TRUE,"",1)'
    book.save(raw)
    assert (
        grade_workbook(
            raw, raw, raw, {"answer_position": "A1"}, evaluated=(paths[1], paths[2])
        )["status"]
        == "pass"
    )


def test_answer_region_ignores_incidental_golden_reference_edits(
    tmp_path: Path,
) -> None:
    from openpyxl.styles import Font

    paths = [tmp_path / name for name in ("input.xlsx", "golden.xlsx", "output.xlsx")]
    book = Workbook()
    book.active["C1"] = "unrelated reference"
    book.save(paths[0])
    book.active["B1"] = "answer"
    book.active["B1"].font = Font(name="Arial")
    book.active["C1"] = None
    book.save(paths[1])
    book.active["C1"] = "unrelated reference"
    book.active["B1"].font = Font(name="Calibri")
    book.save(paths[2])
    assert grade_workbook(*paths, {"answer_position": "B1"})["status"] == "pass"


def test_unchanged_array_formulas_must_be_preserved(tmp_path: Path) -> None:
    from openpyxl.worksheet.formula import ArrayFormula

    paths = [tmp_path / name for name in ("input.xlsx", "golden.xlsx", "output.xlsx")]
    book = Workbook()
    book.active["A1"] = ArrayFormula(ref="A1", text="=SUM(1,2)")
    book.save(paths[0])
    book.active["B1"] = 2
    book.save(paths[1])
    book.active["A1"] = "=42"
    book.save(paths[2])
    result = grade_workbook(*paths, {"answer_position": "B1"})
    assert result["status"] == "incorrect"
    assert result["preservationFailures"] == [{"sheet": "Sheet", "cell": "A1"}]


def test_formula_request_cannot_pass_with_only_computed_values(tmp_path: Path) -> None:
    initial, golden, actual, cached = [
        tmp_path / name
        for name in ("input.xlsx", "golden.xlsx", "output.xlsx", "cached.xlsx")
    ]
    book = Workbook()
    book.active["A1"] = 0
    book.save(initial)
    book.active["B1"] = "=1+2"
    book.save(golden)
    book.active["B1"] = 3
    book.save(actual)
    book.save(cached)
    result = grade_workbook(
        initial,
        golden,
        actual,
        {"answer_position": "B1", "instruction": "Write a formula in B1"},
        evaluated=(cached, actual),
    )
    assert result["status"] == "incorrect"
    assert any(d.get("kind") == "missing_formula" for d in result["valueDifferences"])
