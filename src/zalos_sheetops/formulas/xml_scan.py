"""Streaming formula and error-token scan on the shared expat engine.

Reuses the private worksheet scanner base from zalos_sheetops.io.xlsx_xml so
formula scans stream cells through expat callbacks instead of materializing
ElementTree elements per cell.
"""

import xml.parsers.expat
from pathlib import Path
from zipfile import ZipFile

from zalos_sheetops.io.xlsx_xml import (
    WorksheetSource,
    XlsxWorkbook,
    _parse_cell_range,
    _parse_coordinate,
    _WorksheetContentScanner,
    translate_shared_formula,
)


class XmlFormulaScanUnsupported(Exception):
    def __init__(self, details: dict[str, object]) -> None:
        self.details = details
        super().__init__("XML formula scan cannot handle this workbook feature")


def scan_xlsx_formulas_xml(
    path: Path,
    sheet_name: str | None,
    cell_range: str | None,
    max_results: int | None,
    error_tokens: set[str],
    include: list[str],
    fail_fast: bool,
) -> dict[str, object]:
    bounds = _parse_cell_range(cell_range) if cell_range is not None else None
    include_formulas = "formulas" in include
    include_errors = "errors" in include
    with XlsxWorkbook(path) as workbook:
        if sheet_name is not None:
            worksheets = [workbook.worksheet(sheet_name)]
        else:
            worksheets = workbook.worksheets
        shared_strings = workbook.shared_strings() if include_errors else []
        shared_string_error_indexes = _shared_string_error_indexes(
            shared_strings,
            error_tokens,
        )
        formulas: list[dict[str, str]] = []
        error_items: list[dict[str, str]] = []
        formula_count = 0
        error_count = 0
        stopped_early = False
        for worksheet in worksheets:
            try:
                has_markers = _worksheet_has_scan_markers(
                    workbook.archive,
                    worksheet.path,
                    error_tokens,
                    shared_string_error_indexes,
                    include_formulas,
                    include_errors,
                )
            except KeyError as error:
                raise _invalid_workbook_for_archive(path) from error
            if not has_markers:
                continue
            result = _scan_worksheet(
                workbook,
                worksheet,
                bounds,
                max_results,
                error_tokens,
                shared_strings,
                include_formulas,
                include_errors,
                fail_fast,
                formulas,
                error_items,
            )
            formula_count += int(result["formulaCount"])
            error_count += int(result["errorTokenCount"])
            if result["stoppedEarly"]:
                stopped_early = True
                break

    return {
        "formulaCount": formula_count,
        "errorTokenCount": error_count,
        "formulas": formulas,
        "errorTokens": error_items,
        "stoppedEarly": stopped_early,
    }


def _shared_string_error_indexes(
    shared_strings: list[str],
    error_tokens: set[str],
) -> set[int]:
    return {
        index for index, value in enumerate(shared_strings) if value in error_tokens
    }


def _worksheet_has_scan_markers(
    archive: ZipFile,
    worksheet_path: str,
    error_tokens: set[str],
    shared_string_error_indexes: set[int],
    include_formulas: bool,
    include_errors: bool,
) -> bool:
    formula_markers = [b"<f", b":f"] if include_formulas else []
    error_markers = (
        [token.encode("utf-8") for token in error_tokens] if include_errors else []
    )
    shared_string_markers = (
        [b't="s"', b"t='s'"] if include_errors and shared_string_error_indexes else []
    )
    markers = formula_markers + error_markers + shared_string_markers
    if not markers:
        return False
    overlap = b""
    with archive.open(worksheet_path) as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if chunk == b"":
                return False
            haystack = overlap + chunk
            if any(marker in haystack for marker in markers):
                return True
            overlap = haystack[-64:]


def _scan_worksheet(
    workbook: XlsxWorkbook,
    worksheet: WorksheetSource,
    bounds: tuple[int, int, int, int] | None,
    max_results: int | None,
    error_tokens: set[str],
    shared_strings: list[str],
    include_formulas: bool,
    include_errors: bool,
    fail_fast: bool,
    formulas_out: list[dict[str, str]],
    errors_out: list[dict[str, str]],
) -> dict[str, object]:
    parser = xml.parsers.expat.ParserCreate()
    parser.buffer_text = True
    scanner = _FormulaErrorScanner(
        parser,
        shared_strings,
        sheet_name=worksheet.name,
        bounds=bounds,
        max_results=max_results,
        error_tokens=error_tokens,
        include_formulas=include_formulas,
        include_errors=include_errors,
        fail_fast=fail_fast,
        # The caller's accumulators are appended to directly so maxResults
        # caps returned items across the whole workbook, matching the
        # openpyxl engine.
        formulas_out=formulas_out,
        errors_out=errors_out,
    )
    workbook._parse_worksheet(worksheet, parser, scanner)
    return {
        "formulaCount": scanner.formula_count,
        "errorTokenCount": scanner.error_count,
        "stoppedEarly": scanner.stopped_early,
    }


class _FormulaErrorScanner(_WorksheetContentScanner):
    """Collect formula cells and error-token cells for one worksheet."""

    def __init__(
        self,
        parser: xml.parsers.expat.XMLParserType,
        shared_strings: list[str],
        *,
        sheet_name: str,
        bounds: tuple[int, int, int, int] | None,
        max_results: int | None,
        error_tokens: set[str],
        include_formulas: bool,
        include_errors: bool,
        fail_fast: bool,
        formulas_out: list[dict[str, str]],
        errors_out: list[dict[str, str]],
    ) -> None:
        super().__init__(parser, shared_strings)
        # Formula text capture is only needed when formulas are requested.
        self.capture_formulas = include_formulas
        self._sheet_name = sheet_name
        self._bounds = bounds
        self._max_results = max_results
        self._error_tokens = error_tokens
        self._include_formulas = include_formulas
        self._include_errors = include_errors
        self._fail_fast = fail_fast
        self._formulas_out = formulas_out
        self._errors_out = errors_out
        self._shared_masters: dict[str, tuple[str, str]] = {}
        self.formula_count = 0
        self.error_count = 0
        self.stopped_early = False

    def _row_start(self, attrs: dict[str, str]) -> None:
        if self._bounds is None:
            return
        row_attr = attrs.get("r")
        if row_attr is not None and int(row_attr) > self._bounds[3]:
            self.done = True

    def _cell_end(self) -> None:
        # `done` stops chunk feeding between reads; cells that were already
        # buffered in the current chunk still fire, so drop them here.
        if self.done:
            return
        coordinate = self._coordinate
        if coordinate is None:
            return
        formula_text: str | None = None
        if self._include_formulas and self._formula_parts is not None:
            formula_text = "".join(self._formula_parts)
            # Record masters before any range filtering so followers inside
            # the range can be translated from a master outside it.
            if self._formula_si is not None and formula_text:
                self._shared_masters.setdefault(
                    self._formula_si,
                    (coordinate, formula_text),
                )
        if self._bounds is not None:
            min_col, min_row, max_col, max_row = self._bounds
            row, column = _parse_coordinate(coordinate)
            if row > max_row:
                self.done = True
                return
            if not (min_row <= row and min_col <= column <= max_col):
                return
        if formula_text is not None:
            if not formula_text and (
                self._formula_si is not None or self._formula_type == "shared"
            ):
                master = (
                    self._shared_masters.get(self._formula_si)
                    if self._formula_si is not None
                    else None
                )
                if master is None:
                    raise XmlFormulaScanUnsupported(
                        {
                            "engine": "xml",
                            "feature": "shared_formula_without_text",
                            "sheet": self._sheet_name,
                            "cell": coordinate,
                        }
                    )
                self.formula_count += 1
                if (
                    self._max_results is None
                    or len(self._formulas_out) < self._max_results
                ):
                    self._formulas_out.append(
                        {
                            "sheet": self._sheet_name,
                            "cell": coordinate,
                            "formula": translate_shared_formula(
                                master[0],
                                master[1],
                                coordinate,
                            ),
                        }
                    )
            else:
                self.formula_count += 1
                if (
                    self._max_results is None
                    or len(self._formulas_out) < self._max_results
                ):
                    self._formulas_out.append(
                        {
                            "sheet": self._sheet_name,
                            "cell": coordinate,
                            "formula": (
                                formula_text
                                if formula_text.startswith("=")
                                else f"={formula_text}"
                            ),
                        }
                    )
        if self._include_errors:
            value = self._cell_value()
            if value in self._error_tokens:
                self.error_count += 1
                if (
                    self._max_results is None
                    or len(self._errors_out) < self._max_results
                ):
                    self._errors_out.append(
                        {
                            "sheet": self._sheet_name,
                            "cell": coordinate,
                            "value": value,
                        }
                    )
                if self._fail_fast:
                    self.stopped_early = True
                    self.done = True

    def _row_end(self) -> None:
        pass


def _invalid_workbook_for_archive(path: Path) -> Exception:
    from zalos_sheetops.errors import SheetOpsError

    return SheetOpsError(
        "invalid_workbook",
        "Workbook could not be opened",
        {"path": str(path)},
    )
