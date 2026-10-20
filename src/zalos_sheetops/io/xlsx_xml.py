import posixpath
import re
import xml.etree.ElementTree as ET
import xml.parsers.expat
from collections.abc import Iterator
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.cellref import (
    BUILTIN_FORMATS,
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)
from zalos_sheetops.io.typed_values import TypedValues

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_RELATIONSHIP_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING_ANCHOR_TYPES = frozenset({"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"})
TAG_COUNT_RE = re.compile(rb"<(?:[A-Za-z0-9_]+:)?(f|mergeCell|hyperlink|drawing)[\s>/]")
METADATA_TAG_RE = re.compile(
    rb"<(?:[A-Za-z0-9_]+:)?(?:mergeCell|drawing)(?:\s[^>]*)?/?>"
)
MERGE_CELL_OPEN_RE = re.compile(rb"<(?:[A-Za-z0-9_]+:)?mergeCell(?:\s|>|/)")
DRAWING_OPEN_RE = re.compile(rb"<(?:[A-Za-z0-9_]+:)?drawing(?:\s|>|/)")
REF_ATTRIBUTE_RE = re.compile(rb"""\bref\s*=\s*["']([^"']+)["']""")
RELATIONSHIP_ID_ATTRIBUTE_RE = re.compile(
    rb"""\b(?:[A-Za-z0-9_]+:)?id\s*=\s*["']([^"']+)["']"""
)

_CHUNK_SIZE = 1 << 18

# Cache mapping expat element names (which keep their namespace prefix) to
# local names. The tag vocabulary of a worksheet is tiny, so this stays small.
_LOCAL_NAMES: dict[str, str] = {}


# Plain __slots__ classes instead of dataclasses: RowData is built once per
# streamed row and importing dataclasses costs ~6ms of CLI startup.
class WorksheetSource:
    __slots__ = ("name", "path", "hidden", "state", "position")

    def __init__(
        self,
        name: str,
        path: str,
        hidden: bool,
        state: str,
        position: int,
    ) -> None:
        self.name = name
        self.path = path
        self.hidden = hidden
        self.state = state
        self.position = position


class CellData:
    __slots__ = (
        "address",
        "row",
        "column",
        "value",
        "formula",
        "cached_value",
        "data_type",
        "number_format",
        "source_value",
    )

    def __init__(
        self,
        address: str,
        row: int,
        column: int,
        value: object,
        formula: str | None,
        cached_value: object,
        data_type: str,
        number_format: str | None,
        source_value: object = None,
    ) -> None:
        self.address = address
        self.row = row
        self.column = column
        self.value = value
        self.formula = formula
        self.cached_value = cached_value
        self.data_type = data_type
        self.number_format = number_format
        self.source_value = source_value


class RowData:
    __slots__ = ("number", "values")

    def __init__(self, number: int, values: list[object]) -> None:
        self.number = number
        self.values = values


class XlsxWorkbook:
    """An open workbook archive with cached metadata and shared strings.

    Reusing one instance across scans avoids re-reading the zip central
    directory and re-parsing the shared strings table for every pass.
    """

    def __init__(self, path: Path, value_mode: str = "raw") -> None:
        from zalos_sheetops.io.typed_values import validate_value_mode

        validate_value_mode(value_mode)
        self._value_mode = value_mode
        self._typed_values = None
        self.path = path
        try:
            self._archive = ZipFile(path, "r")
        except (BadZipFile, OSError) as error:
            raise _invalid_workbook(path) from error
        try:
            self.worksheets = _resolve_worksheets(self._archive, path)
        except (ET.ParseError, KeyError) as error:
            self._archive.close()
            raise _invalid_workbook(path) from error
        except SheetOpsError:
            self._archive.close()
            raise
        self._shared_strings: list[str] | None = None
        self._number_formats: dict[int, str] | None = None

    def __enter__(self) -> "XlsxWorkbook":
        return self

    def __exit__(self, *exc_info: object) -> None:
        self.close()

    def close(self) -> None:
        self._archive.close()

    @property
    def archive(self) -> ZipFile:
        return self._archive

    def sheet_names(self) -> list[str]:
        return [worksheet.name for worksheet in self.worksheets]

    def worksheet(self, sheet_name: str) -> WorksheetSource:
        for worksheet in self.worksheets:
            if worksheet.name == sheet_name:
                return worksheet
        raise SheetOpsError(
            "invalid_sheet",
            "Requested sheet does not exist",
            {"sheet": sheet_name},
        )

    def shared_strings(self) -> list[str]:
        if self._shared_strings is None:
            try:
                self._shared_strings = _read_shared_strings(self._archive)
            except (xml.parsers.expat.ExpatError, ValueError) as error:
                raise _invalid_workbook(self.path) from error
        return self._shared_strings

    def number_formats(self) -> dict[int, str]:
        if self._number_formats is None:
            try:
                self._number_formats = _read_number_formats(self._archive)
            except (ET.ParseError, ValueError) as error:
                raise _invalid_workbook(self.path) from error
        return self._number_formats

    def iter_rows(self, sheet_name: str) -> Iterator[RowData]:
        worksheet = self.worksheet(sheet_name)
        shared_strings = self.shared_strings()
        rows: list[RowData] = []
        parser = xml.parsers.expat.ParserCreate()
        parser.buffer_text = True
        _RowScanner(parser, shared_strings, rows, self.typed_values())
        try:
            with self._archive.open(worksheet.path) as handle:
                while True:
                    chunk = handle.read(_CHUNK_SIZE)
                    if not chunk:
                        parser.Parse(b"", True)
                        break
                    parser.Parse(chunk, False)
                    if rows:
                        yield from rows
                        rows.clear()
        except (xml.parsers.expat.ExpatError, KeyError, ValueError) as error:
            raise _invalid_workbook(self.path) from error
        if rows:
            yield from rows

    def scan_cells(
        self,
        sheet_name: str,
        *,
        bounds: tuple[int, int, int, int] | None = None,
        columns: set[int] | None = None,
    ) -> list[CellData]:
        """Collect CellData for cells inside bounds or the given columns.

        A row bound allows the scan to stop before parsing the rest of the
        worksheet.
        """
        worksheet = self.worksheet(sheet_name)
        cells: list[CellData] = []
        parser = xml.parsers.expat.ParserCreate()
        parser.buffer_text = True
        scanner = _CellScanner(
            parser,
            self.shared_strings(),
            self.number_formats(),
            cells,
            bounds=bounds,
            columns=columns,
        )
        self._parse_worksheet(worksheet, parser, scanner)
        converter = self.typed_values()
        if converter is not None:
            from zalos_sheetops.io.typed_values import json_typed_value

            # scan_cells already resolves number formats. Convert only numeric
            # date cells and ISO date storage; formulas keep their original text.
            styles = self.number_formats()
            format_styles = {fmt: style for style, fmt in styles.items()}
            for cell in cells:
                cell.value = json_typed_value(
                    converter.convert(
                        cell.value,
                        str(format_styles.get(cell.number_format, 0)),
                        cell.data_type,
                    )
                )
                cell.cached_value = cell.value
        return cells

    def typed_values(self) -> TypedValues | None:
        if self._value_mode == "raw":
            return None
        if self._typed_values is None:
            root = ET.fromstring(self._archive.read("xl/workbook.xml"))
            properties = root.find(f"{{{MAIN_NS}}}workbookPr")
            date_1904 = properties is not None and properties.get("date1904") in (
                "1",
                "true",
            )
            self._typed_values = TypedValues(self.number_formats(), date_1904)
        return self._typed_values

    def value_metadata(self) -> dict[str, str]:
        converter = self.typed_values()
        if converter is None:
            return {}
        return {"valueMode": "typed", "dateSystem": converter.date_system}

    def scan_formula_cells(self, sheet_name: str, max_cells: int) -> list[str]:
        worksheet = self.worksheet(sheet_name)
        coordinates: list[str] = []
        parser = xml.parsers.expat.ParserCreate()
        parser.buffer_text = True
        scanner = _FormulaCellScanner(parser, coordinates, max_cells)
        self._parse_worksheet(worksheet, parser, scanner)
        return coordinates

    def _parse_worksheet(
        self,
        worksheet: WorksheetSource,
        parser: xml.parsers.expat.XMLParserType,
        scanner: object,
    ) -> None:
        try:
            with self._archive.open(worksheet.path) as handle:
                while True:
                    chunk = handle.read(_CHUNK_SIZE)
                    if not chunk:
                        parser.Parse(b"", True)
                        break
                    parser.Parse(chunk, False)
                    if scanner.done:  # type: ignore[attr-defined]
                        break
        except (xml.parsers.expat.ExpatError, KeyError, ValueError) as error:
            raise _invalid_workbook(self.path) from error


_DIGITS = "0123456789"

# Column letters repeat on every row, so cache their parsed index.
_COLUMN_INDEX_CACHE: dict[str, int] = {}

# openpyxl's formula translator, imported on first shared-formula follower.
_TRANSLATOR: type | None = None


def translate_shared_formula(origin: str, master_text: str, coordinate: str) -> str:
    """Derive a shared-formula follower's formula from its master cell.

    Uses openpyxl's Translator (the same mechanism openpyxl applies when it
    expands shared formulas), imported lazily so workbooks without shared
    formulas never pay the openpyxl import.
    """
    global _TRANSLATOR
    if _TRANSLATOR is None:
        from openpyxl.formula.translate import Translator

        _TRANSLATOR = Translator
    return _TRANSLATOR(f"={master_text}", origin=origin).translate_formula(coordinate)


def _column_letters_index(letters: str) -> int:
    column = 0
    for character in letters:
        code = ord(character)
        if 65 <= code <= 90:
            column = column * 26 + code - 64
        elif 97 <= code <= 122:
            column = column * 26 + code - 96
        else:
            raise ValueError(f"Invalid column letters {letters}")
    if not column:
        raise ValueError(f"Invalid column letters {letters}")
    _COLUMN_INDEX_CACHE[letters] = column
    return column


def _column_of(coordinate: str) -> int:
    letters = coordinate.rstrip(_DIGITS)
    column = _COLUMN_INDEX_CACHE.get(letters)
    if column is None:
        column = _column_letters_index(letters)
    return column


def _parse_coordinate(coordinate: str) -> tuple[int, int]:
    letters = coordinate.rstrip(_DIGITS)
    column = _COLUMN_INDEX_CACHE.get(letters)
    if column is None:
        column = _column_letters_index(letters)
    return int(coordinate[len(letters) :]), column


def _numeric_or_text(value: str) -> object:
    try:
        return int(value)
    except ValueError:
        try:
            return float(value)
        except ValueError:
            return value


class _WorksheetContentScanner:
    """Shared expat handling of worksheet cell content.

    Tracks the current cell coordinate, type, value text, inline string, and
    formula. Subclasses decide what to do when a cell or row ends.

    Element names are compared directly against the unprefixed spellings
    first: expat interns element names, so those comparisons are pointer
    checks, and virtually all real workbooks use a default namespace. The
    fallback resolves prefixed names through the shared cache.
    """

    capture_formulas = False

    def __init__(
        self,
        parser: xml.parsers.expat.XMLParserType,
        shared_strings: list[str],
    ) -> None:
        self._parser = parser
        self._shared_strings = shared_strings
        self._coordinate: str | None = None
        self._cell_type: str | None = None
        self._style: str | None = None
        self._value_text: str | None = None
        self._inline_parts: list[str] | None = None
        self._formula_parts: list[str] | None = None
        self._formula_type: str | None = None
        self._formula_si: str | None = None
        self._capture: list[str] | None = None
        self._row_attr: str | None = None
        self.done = False
        parser.StartElementHandler = self._start
        parser.EndElementHandler = self._end

    def _start(self, name: str, attrs: dict[str, str]) -> None:
        if name == "c":
            self._coordinate = attrs.get("r")
            self._cell_type = attrs.get("t")
            self._style = attrs.get("s")
            self._value_text = None
            self._inline_parts = None
            self._formula_parts = None
            self._formula_type = None
            self._formula_si = None
        elif name == "v":
            self._capture = []
            self._parser.CharacterDataHandler = self._capture.append
        elif name == "f":
            self._formula_type = attrs.get("t")
            self._formula_si = attrs.get("si")
            if self.capture_formulas:
                self._formula_parts = []
                self._capture = self._formula_parts
                self._parser.CharacterDataHandler = self._formula_parts.append
        elif name == "row":
            self._row_start(attrs)
        else:
            local = _LOCAL_NAMES.get(name)
            if local is None:
                local = name.rpartition(":")[2]
                _LOCAL_NAMES[name] = local
            if local == "is":
                self._inline_parts = []
            elif local == "t":
                if self._inline_parts is not None:
                    self._capture = []
                    self._parser.CharacterDataHandler = self._capture.append
            elif local != name:
                self._start(local, attrs)

    def _end(self, name: str) -> None:
        if name == "c":
            self._cell_end()
        elif name == "v":
            capture = self._capture
            if capture is not None:
                text = "".join(capture)
                self._value_text = text if text else None
                self._capture = None
                self._parser.CharacterDataHandler = None
        elif name == "f":
            if self._formula_parts is not None and (
                self._capture is self._formula_parts
            ):
                self._capture = None
                self._parser.CharacterDataHandler = None
        elif name == "row":
            self._row_end()
        else:
            local = _LOCAL_NAMES.get(name)
            if local is None:
                local = name.rpartition(":")[2]
                _LOCAL_NAMES[name] = local
            if local == "t":
                capture = self._capture
                if capture is not None and self._inline_parts is not None:
                    if capture is not self._inline_parts:
                        self._inline_parts.append("".join(capture))
                    self._capture = None
                    self._parser.CharacterDataHandler = None
            elif local != name:
                self._end(local)

    def _cell_value(self) -> object:
        cell_type = self._cell_type
        if cell_type is None or cell_type == "n":
            text = self._value_text
            if text is None:
                return None
            try:
                return int(text)
            except ValueError:
                try:
                    return float(text)
                except ValueError:
                    return text
        if cell_type == "inlineStr":
            parts = self._inline_parts
            return "".join(parts) if parts else ""
        text = self._value_text
        if text is None:
            return None
        if cell_type == "s":
            try:
                index = int(text)
            except ValueError:
                return None
            shared_strings = self._shared_strings
            if 0 <= index < len(shared_strings):
                return shared_strings[index]
            return None
        if cell_type == "b":
            return text == "1"
        if cell_type == "e":
            return text
        return _numeric_or_text(text)

    def _row_start(self, attrs: dict[str, str]) -> None:
        raise NotImplementedError

    def _cell_end(self) -> None:
        raise NotImplementedError

    def _row_end(self) -> None:
        raise NotImplementedError


class _RowScanner(_WorksheetContentScanner):
    """Collect one RowData per worksheet row.

    This is the hottest scanner (read-table, validate, and normalize all
    stream every row), so its handlers are specialized: no style attribute
    read, formula elements ignored, and the column parse and numeric value
    conversion inlined into the cell-end handler.
    """

    def __init__(
        self,
        parser: xml.parsers.expat.XMLParserType,
        shared_strings: list[str],
        rows_out: list[RowData],
        typed_values: TypedValues | None = None,
    ) -> None:
        super().__init__(parser, shared_strings)
        self._rows_out = rows_out
        self._typed_values = typed_values
        self._values: list[object] | None = None
        self._next_column = 1
        self._first_cell_row: int | None = None

    def _start(self, name: str, attrs: dict[str, str]) -> None:
        if name == "c":
            self._coordinate = attrs.get("r")
            self._cell_type = attrs.get("t")
            if self._typed_values is not None:
                self._style = attrs.get("s")
            self._value_text = None
            self._inline_parts = None
        elif name == "v":
            self._capture = []
            self._parser.CharacterDataHandler = self._capture.append
        elif name == "t":
            # Hoisted above the namespace fallback: inline-string sheets
            # (openpyxl write-only output) hit this branch once per cell.
            if self._inline_parts is not None:
                self._capture = []
                self._parser.CharacterDataHandler = self._capture.append
        elif name == "is":
            self._inline_parts = []
        elif name == "f":
            pass
        elif name == "row":
            self._row_start(attrs)
        else:
            local = _LOCAL_NAMES.get(name)
            if local is None:
                local = name.rpartition(":")[2]
                _LOCAL_NAMES[name] = local
            if local != name:
                self._start(local, attrs)

    def _end(self, name: str) -> None:
        if name == "c":
            values = self._values
            if values is None:
                return
            coordinate = self._coordinate
            if coordinate is None:
                column = self._next_column
            else:
                letters = coordinate.rstrip(_DIGITS)
                column = _COLUMN_INDEX_CACHE.get(letters)
                if column is None:
                    column = _column_letters_index(letters)
                if self._row_attr is None and self._first_cell_row is None:
                    self._first_cell_row = int(coordinate[len(letters) :])
            self._next_column = column + 1
            if self._cell_type is None:
                text = self._value_text
                if text is None:
                    value = None
                else:
                    try:
                        value = int(text)
                    except ValueError:
                        try:
                            value = float(text)
                        except ValueError:
                            value = text
            else:
                value = self._cell_value()
            if self._typed_values is not None:
                value = self._typed_values.convert(value, self._style, self._cell_type)
            index = column - 1
            if index == len(values):
                values.append(value)
            elif index < len(values):
                values[index] = value
            else:
                values.extend([None] * (index - len(values)))
                values.append(value)
        elif name == "v":
            capture = self._capture
            if capture is not None:
                text = "".join(capture)
                self._value_text = text if text else None
                self._capture = None
                self._parser.CharacterDataHandler = None
        elif name == "t":
            capture = self._capture
            if capture is not None and self._inline_parts is not None:
                if capture is not self._inline_parts:
                    self._inline_parts.append("".join(capture))
                self._capture = None
                self._parser.CharacterDataHandler = None
        elif name == "f":
            pass
        elif name == "row":
            self._row_end()
        else:
            local = _LOCAL_NAMES.get(name)
            if local is None:
                local = name.rpartition(":")[2]
                _LOCAL_NAMES[name] = local
            if local != name:
                self._end(local)

    def _row_start(self, attrs: dict[str, str]) -> None:
        self._row_attr = attrs.get("r")
        self._values = []
        self._next_column = 1
        self._first_cell_row = None

    def _cell_end(self) -> None:  # pragma: no cover - handled inline in _end
        raise NotImplementedError

    def _row_end(self) -> None:
        values = self._values
        if values is None:
            return
        row_attr = self._row_attr
        if row_attr is not None:
            number = int(row_attr)
        elif self._first_cell_row is not None:
            number = self._first_cell_row
        else:
            number = 0
        self._rows_out.append(RowData(number=number, values=values))
        self._values = None


class _CellScanner(_WorksheetContentScanner):
    """Collect CellData for cells matching bounds or a column filter."""

    capture_formulas = True

    def __init__(
        self,
        parser: xml.parsers.expat.XMLParserType,
        shared_strings: list[str],
        number_formats: dict[int, str],
        cells_out: list[CellData],
        *,
        bounds: tuple[int, int, int, int] | None = None,
        columns: set[int] | None = None,
    ) -> None:
        super().__init__(parser, shared_strings)
        self._number_formats = number_formats
        self._cells_out = cells_out
        self._bounds = bounds
        self._columns = columns
        self._current_row = 0
        self._shared_masters: dict[str, tuple[str, str]] = {}

    def _row_start(self, attrs: dict[str, str]) -> None:
        row_attr = attrs.get("r")
        if row_attr is not None:
            self._current_row = int(row_attr)
            if self._bounds is not None and self._current_row > self._bounds[3]:
                self.done = True

    def _cell_end(self) -> None:
        coordinate = self._coordinate
        if coordinate is None:
            return
        formula_parts = self._formula_parts
        formula_text: str | None = None
        if formula_parts is not None:
            formula_text = "".join(formula_parts)
            # Shared-formula masters must be recorded even when this cell is
            # outside the requested bounds, so in-range followers can be
            # translated.
            if self._formula_si is not None and formula_text:
                self._shared_masters.setdefault(
                    self._formula_si,
                    (coordinate, formula_text),
                )
        row, column = _parse_coordinate(coordinate)
        bounds = self._bounds
        if bounds is not None:
            min_col, min_row, max_col, max_row = bounds
            if row > max_row:
                self.done = True
                return
            if not (min_row <= row and min_col <= column <= max_col):
                return
        if self._columns is not None and column not in self._columns:
            return
        formula = None
        if formula_text is not None:
            if self._formula_si is not None and not formula_text:
                master = self._shared_masters.get(self._formula_si)
                if master is not None:
                    formula = translate_shared_formula(
                        master[0],
                        master[1],
                        coordinate,
                    )
                else:
                    formula = "="
            else:
                formula = (
                    formula_text if formula_text.startswith("=") else f"={formula_text}"
                )
        value = self._cell_value()
        style_id = 0
        style = self._style
        if style is not None:
            try:
                style_id = int(style)
            except ValueError:
                style_id = 0
        self._cells_out.append(
            CellData(
                address=coordinate,
                row=row,
                column=column,
                value=value,
                formula=formula,
                cached_value=value,
                data_type=self._cell_type or "n",
                number_format=self._number_formats.get(style_id, "General"),
                source_value=(
                    self._value_text
                    if self._cell_type in {None, "n", "str", "d", "e"}
                    else value
                ),
            )
        )

    def _row_end(self) -> None:
        pass


class _FormulaCellScanner:
    """Collect coordinates of cells that contain formulas, up to a limit."""

    def __init__(
        self,
        parser: xml.parsers.expat.XMLParserType,
        coordinates_out: list[str],
        max_cells: int,
    ) -> None:
        self._coordinates_out = coordinates_out
        self._max_cells = max_cells
        self._coordinate: str | None = None
        self._has_formula = False
        self.done = max_cells <= 0
        parser.StartElementHandler = self._start
        parser.EndElementHandler = self._end

    def _start(self, name: str, attrs: dict[str, str]) -> None:
        if name == "c":
            self._coordinate = attrs.get("r")
            self._has_formula = False
        elif name == "f":
            self._has_formula = True
        else:
            local = _LOCAL_NAMES.get(name)
            if local is None:
                local = name.rpartition(":")[2]
                _LOCAL_NAMES[name] = local
            if local != name:
                self._start(local, attrs)

    def _end(self, name: str) -> None:
        if name != "c":
            local = _LOCAL_NAMES.get(name)
            if local is None:
                local = name.rpartition(":")[2]
                _LOCAL_NAMES[name] = local
            if local != "c":
                return
        if not self.done and self._has_formula and self._coordinate is not None:
            self._coordinates_out.append(self._coordinate)
            if len(self._coordinates_out) >= self._max_cells:
                self.done = True


def read_workbook_sources(path: Path) -> list[WorksheetSource]:
    try:
        with ZipFile(path, "r") as archive:
            return _resolve_worksheets(archive, path)
    except (BadZipFile, ET.ParseError, KeyError) as error:
        raise _invalid_workbook(path) from error


def read_workbook_xml_summary(
    path: Path,
    *,
    max_formula_cells: int,
    max_merged_ranges: int = 1000,
    max_graphics: int = 1000,
    max_defined_names: int = 1000,
) -> dict[str, object]:
    try:
        with XlsxWorkbook(path) as workbook:
            archive = workbook.archive
            sheets = []
            formula_count = 0
            for worksheet in workbook.worksheets:
                sheet_summary = _read_sheet_xml_summary(
                    workbook,
                    archive,
                    worksheet,
                    max_formula_cells,
                    max_merged_ranges,
                    max_graphics,
                    path,
                )
                formula_count += int(sheet_summary["formulaCount"])
                sheets.append(sheet_summary)
            defined_name_count, defined_names = _read_defined_names(
                archive,
                workbook.worksheets,
                max_defined_names,
                path,
            )
            return {
                "workbook": {
                    "path": str(path),
                    "sheetCount": len(workbook.worksheets),
                    "formulaCount": formula_count,
                    "engine": "xml",
                    "definedNameCount": defined_name_count,
                    "definedNames": defined_names,
                    "definedNamesTruncated": defined_name_count > len(defined_names),
                },
                "sheets": sheets,
            }
    except (BadZipFile, ET.ParseError, KeyError) as error:
        raise _invalid_workbook(path) from error


def read_cells_xml(
    path: Path,
    sheet_name: str,
    cell_range: str,
    include_fields: list[str],
    value_mode: str = "raw",
) -> dict[str, object]:
    min_col, min_row, max_col, max_row = _parse_cell_range(cell_range)
    with XlsxWorkbook(path, value_mode=value_mode) as workbook:
        parsed = workbook.scan_cells(
            sheet_name,
            bounds=(min_col, min_row, max_col, max_row),
        )
        value_metadata = workbook.value_metadata()

    # Last write wins for duplicate addresses, matching the previous
    # grid-lookup behavior.
    parsed_cells = {cell.address: cell for cell in parsed}
    cells = []
    for cell in parsed_cells.values():
        if cell.value is None and cell.formula is None:
            continue
        cells.append(_cell_payload(cell, include_fields))

    return {
        "engine": "xml",
        "sheet": sheet_name,
        "range": cell_range,
        "rowCount": max_row - min_row + 1,
        "columnCount": max_col - min_col + 1,
        "cellCount": (max_row - min_row + 1) * (max_col - min_col + 1),
        "returnedCellCount": len(cells),
        "cells": cells,
        **value_metadata,
    }


def read_labeled_cells_xml(
    path: Path,
    sheet_name: str,
    label_column: str,
    value_column: str,
    requested_labels: set[str] | None = None,
) -> dict[str, CellData]:
    label_index = _column_letter_to_index(label_column, "labelColumn")
    value_index = _column_letter_to_index(value_column, "valueColumn")
    with XlsxWorkbook(path) as workbook:
        parsed = workbook.scan_cells(
            sheet_name,
            columns={label_index, value_index},
        )

    row_cells: dict[int, dict[int, CellData]] = {}
    for cell in parsed:
        row_cells.setdefault(cell.row, {})[cell.column] = cell

    labels: dict[str, CellData] = {}
    for row_number, cells in row_cells.items():
        label_cell = cells.get(label_index)
        if label_cell is None or not isinstance(label_cell.value, str):
            continue
        if requested_labels is not None and label_cell.value not in requested_labels:
            continue
        if label_cell.value in labels:
            raise SheetOpsError(
                "duplicate_label",
                "Label appears more than once in sheet",
                {"sheet": sheet_name, "label": label_cell.value},
            )
        value_cell = cells.get(value_index)
        if value_cell is None:
            address = f"{get_column_letter(value_index)}{row_number}"
            value_cell = CellData(
                address=address,
                row=row_number,
                column=value_index,
                value=None,
                formula=None,
                cached_value=None,
                data_type="blank",
                number_format=None,
            )
        labels[label_cell.value] = value_cell
    return labels


def iter_sheet_rows_xml(path: Path, sheet_name: str) -> Iterator[RowData]:
    with XlsxWorkbook(path) as workbook:
        yield from workbook.iter_rows(sheet_name)


def _read_sheet_xml_summary(
    workbook: XlsxWorkbook,
    archive: ZipFile,
    worksheet: WorksheetSource,
    max_formula_cells: int,
    max_merged_ranges: int,
    max_graphics: int,
    path: Path,
) -> dict[str, object]:
    dimension = _read_dimension(archive, worksheet.path)
    max_row, max_column = _dimension_size(dimension)
    counts, merged_ranges, drawing_relationship_ids = _count_sheet_tags(
        archive,
        worksheet.path,
        max_merged_ranges,
        path,
    )
    formula_count = counts["formulaCount"]
    formula_cells: list[str] = []
    if formula_count and max_formula_cells > 0:
        formula_cells = workbook.scan_formula_cells(
            worksheet.name,
            max_formula_cells,
        )
    graphic_count, graphics = _read_sheet_graphics(
        archive,
        worksheet,
        drawing_relationship_ids,
        max_graphics,
        path,
    )

    return {
        "name": worksheet.name,
        "maxRow": max_row,
        "maxColumn": max_column,
        "hidden": worksheet.hidden,
        "state": worksheet.state,
        "formulaCount": formula_count,
        "formulaCells": formula_cells,
        "formulaCellsTruncated": formula_count > len(formula_cells),
        "dimensions": dimension,
        "mergedRangeCount": counts["mergedRangeCount"],
        "mergedRanges": merged_ranges,
        "mergedRangesTruncated": counts["mergedRangeCount"] > len(merged_ranges),
        "hyperlinkCount": counts["hyperlinkCount"],
        "graphicCount": graphic_count,
        "graphics": graphics,
        "graphicsTruncated": graphic_count > len(graphics),
    }


def _read_defined_names(
    archive: ZipFile,
    worksheets: list[WorksheetSource],
    max_defined_names: int,
    path: Path,
) -> tuple[int, list[dict[str, object]]]:
    root = _read_xml(archive, "xl/workbook.xml")
    elements = root.findall(f".//{{{MAIN_NS}}}definedName")
    names = []
    for index, element in enumerate(elements):
        name = element.attrib.get("name")
        if name is None:
            raise _invalid_workbook(path)
        scope = _defined_name_scope(element, worksheets, path)
        if index >= max_defined_names:
            continue
        names.append(
            {
                "name": name,
                "value": element.text or "",
                "scope": scope,
                "hidden": _xml_bool(element.attrib.get("hidden")),
                "builtIn": name.startswith("_xlnm."),
            }
        )
    return len(elements), names


def _defined_name_scope(
    element: ET.Element,
    worksheets: list[WorksheetSource],
    path: Path,
) -> str:
    local_sheet_id = element.attrib.get("localSheetId")
    if local_sheet_id is None:
        return "workbook"
    try:
        index = int(local_sheet_id)
    except ValueError as error:
        raise _invalid_workbook(path) from error
    if index < 0 or index >= len(worksheets):
        raise _invalid_workbook(path)
    return worksheets[index].name


def _xml_bool(value: str | None) -> bool:
    return value in {"1", "true", "True"}


def _collect_metadata_tags(
    data: bytes,
    merged_ranges: list[str],
    drawing_relationship_ids: list[str],
    max_merged_ranges: int,
    path: Path,
) -> None:
    for match in METADATA_TAG_RE.finditer(data):
        tag = match.group(0)
        if MERGE_CELL_OPEN_RE.match(tag):
            reference = _required_xml_attribute(tag, REF_ATTRIBUTE_RE, path)
            if len(merged_ranges) < max_merged_ranges:
                merged_ranges.append(reference)
        elif DRAWING_OPEN_RE.match(tag):
            drawing_relationship_ids.append(
                _required_xml_attribute(
                    tag,
                    RELATIONSHIP_ID_ATTRIBUTE_RE,
                    path,
                )
            )


def _required_xml_attribute(
    tag: bytes,
    pattern: re.Pattern[bytes],
    path: Path,
) -> str:
    match = pattern.search(tag)
    if match is None:
        raise _invalid_workbook(path)
    try:
        return match.group(1).decode("utf-8")
    except UnicodeDecodeError as error:
        raise _invalid_workbook(path) from error


def _read_sheet_graphics(
    archive: ZipFile,
    worksheet: WorksheetSource,
    drawing_relationship_ids: list[str],
    max_graphics: int,
    path: Path,
) -> tuple[int, list[dict[str, object]]]:
    if not drawing_relationship_ids:
        return 0, []
    relationships = _read_relationships(
        archive,
        _relationships_path(worksheet.path),
    )
    count = 0
    graphics: list[dict[str, object]] = []
    for relationship_id in drawing_relationship_ids:
        drawing_path = relationships.get(relationship_id)
        if drawing_path is None:
            raise _invalid_workbook(path)
        drawing_count, drawing_graphics = _read_drawing_graphics(
            archive,
            drawing_path,
            max(0, max_graphics - len(graphics)),
            path,
        )
        count += drawing_count
        graphics.extend(drawing_graphics)
    return count, graphics


def _read_drawing_graphics(
    archive: ZipFile,
    drawing_path: str,
    max_graphics: int,
    path: Path,
) -> tuple[int, list[dict[str, object]]]:
    root = _read_xml(archive, drawing_path)
    relationships = _read_relationships(
        archive,
        _relationships_path(drawing_path),
        missing_ok=True,
    )
    count = 0
    graphics: list[dict[str, object]] = []
    for anchor in _iter_drawing_anchors(root):
        anchor_payload = _drawing_anchor_payload(anchor, path)
        graphic = _graphic_payload(
            archive,
            anchor,
            anchor_payload,
            relationships,
            path,
        )
        if graphic is None:
            continue
        count += 1
        if len(graphics) < max_graphics:
            graphics.append(graphic)
    return count, graphics


def _iter_drawing_anchors(element: ET.Element) -> Iterator[ET.Element]:
    local_name = _local_name(element.tag)
    if local_name in DRAWING_ANCHOR_TYPES:
        yield element
        return
    children = list(element)
    if local_name == "AlternateContent":
        selected = next(
            (
                child
                for child in children
                if _local_name(child.tag) == "Choice"
                and _contains_drawing_anchor(child)
            ),
            None,
        )
        if selected is None:
            selected = next(
                (child for child in children if _local_name(child.tag) == "Fallback"),
                None,
            )
        if selected is not None:
            yield from _iter_drawing_anchors(selected)
        return
    for child in children:
        yield from _iter_drawing_anchors(child)


def _contains_drawing_anchor(element: ET.Element) -> bool:
    return any(
        _local_name(descendant.tag) in DRAWING_ANCHOR_TYPES
        for descendant in element.iter()
    )


def _graphic_payload(
    archive: ZipFile,
    anchor: ET.Element,
    anchor_payload: dict[str, object],
    relationships: dict[str, str],
    path: Path,
) -> dict[str, object] | None:
    chart = anchor.find(f".//{{{CHART_NS}}}chart")
    if chart is not None:
        relationship_id = chart.attrib.get(f"{{{RELATIONSHIP_NS}}}id")
        chart_path = _required_relationship_target(
            relationship_id,
            relationships,
            path,
        )
        if chart_path not in archive.namelist():
            raise _invalid_workbook(path)
        return {
            "type": "chart",
            "title": _read_chart_title(archive, chart_path),
            "anchor": anchor_payload,
        }
    picture = anchor.find(f".//{{{DRAWING_NS}}}pic")
    if picture is not None:
        blip = picture.find(f".//{{{DRAWINGML_NS}}}blip")
        if blip is None:
            raise _invalid_workbook(path)
        embedded_relationship_id = blip.attrib.get(f"{{{RELATIONSHIP_NS}}}embed")
        linked_relationship_id = blip.attrib.get(f"{{{RELATIONSHIP_NS}}}link")
        relationship_id = embedded_relationship_id or linked_relationship_id
        image_path = _required_relationship_target(
            relationship_id,
            relationships,
            path,
        )
        if (
            embedded_relationship_id is not None
            and image_path not in archive.namelist()
        ):
            raise _invalid_workbook(path)
        properties = picture.find(f".//{{{DRAWING_NS}}}cNvPr")
        title = None
        if properties is not None:
            title = properties.attrib.get("descr") or properties.attrib.get("name")
        return {"type": "image", "title": title, "anchor": anchor_payload}
    return None


def _required_relationship_target(
    relationship_id: str | None,
    relationships: dict[str, str],
    path: Path,
) -> str:
    if relationship_id is None:
        raise _invalid_workbook(path)
    target = relationships.get(relationship_id)
    if target is None:
        raise _invalid_workbook(path)
    return target


def _drawing_anchor_payload(
    anchor: ET.Element,
    path: Path,
) -> dict[str, object]:
    anchor_type = _local_name(anchor.tag)
    if anchor_type == "twoCellAnchor":
        start = _anchor_marker(anchor, "from", path)
        end = _anchor_marker(anchor, "to", path)
        return {
            "type": "twoCell",
            "from": start,
            "to": end,
            "range": f"{start}:{end}",
        }
    if anchor_type == "oneCellAnchor":
        start = _anchor_marker(anchor, "from", path)
        extent = _extent_payload(anchor, path)
        return {"type": "oneCell", "from": start, "extentEmu": extent}
    if anchor_type == "absoluteAnchor":
        position = _position_payload(anchor, path)
        extent = _extent_payload(anchor, path)
        return {
            "type": "absolute",
            "positionEmu": position,
            "extentEmu": extent,
        }
    raise _invalid_workbook(path)


def _anchor_marker(anchor: ET.Element, marker_name: str, path: Path) -> str:
    marker = anchor.find(f"{{{DRAWING_NS}}}{marker_name}")
    if marker is None:
        raise _invalid_workbook(path)
    column = _required_child_int(marker, "col", path)
    row = _required_child_int(marker, "row", path)
    if column < 0 or row < 0:
        raise _invalid_workbook(path)
    return f"{get_column_letter(column + 1)}{row + 1}"


def _extent_payload(anchor: ET.Element, path: Path) -> dict[str, int]:
    extent = anchor.find(f"{{{DRAWING_NS}}}ext")
    if extent is None:
        raise _invalid_workbook(path)
    return {
        "width": _required_int_attribute(extent, "cx", path),
        "height": _required_int_attribute(extent, "cy", path),
    }


def _position_payload(anchor: ET.Element, path: Path) -> dict[str, int]:
    position = anchor.find(f"{{{DRAWING_NS}}}pos")
    if position is None:
        raise _invalid_workbook(path)
    return {
        "x": _required_int_attribute(position, "x", path),
        "y": _required_int_attribute(position, "y", path),
    }


def _required_child_int(
    element: ET.Element,
    child_name: str,
    path: Path,
) -> int:
    child = element.find(f"{{{DRAWING_NS}}}{child_name}")
    if child is None or child.text is None:
        raise _invalid_workbook(path)
    try:
        return int(child.text)
    except ValueError as error:
        raise _invalid_workbook(path) from error


def _required_int_attribute(
    element: ET.Element,
    attribute: str,
    path: Path,
) -> int:
    value = element.attrib.get(attribute)
    if value is None:
        raise _invalid_workbook(path)
    try:
        return int(value)
    except ValueError as error:
        raise _invalid_workbook(path) from error


def _read_chart_title(archive: ZipFile, chart_path: str | None) -> str | None:
    if chart_path is None:
        return None
    try:
        root = _read_xml(archive, chart_path)
    except KeyError:
        return None
    title = root.find(f".//{{{CHART_NS}}}title")
    if title is None:
        return None
    text = "".join(
        element.text or "" for element in title.iter(f"{{{DRAWINGML_NS}}}t")
    ).strip()
    return text or None


def _count_sheet_tags(
    archive: ZipFile,
    worksheet_path: str,
    max_merged_ranges: int,
    path: Path,
) -> tuple[dict[str, int], list[str], list[str]]:
    formula_count = 0
    merged_range_count = 0
    hyperlink_count = 0
    merged_ranges: list[str] = []
    drawing_relationship_ids: list[str] = []
    find_tags = TAG_COUNT_RE.findall
    tail = b""
    try:
        with archive.open(worksheet_path) as handle:
            while True:
                chunk = handle.read(1024 * 1024)
                if not chunk:
                    break
                data = tail + chunk
                # Split at a tag start inside the final window so no match can
                # straddle the cut: a match never contains a second "<", so
                # every match starting before the split completes inside the
                # segment, and matches starting at or after it are counted
                # with the next chunk. Tags longer than the 128-byte window
                # are not supported (matching the previous allowance).
                split = data.find(b"<", max(0, len(data) - 128))
                if split == -1:
                    split = len(data)
                segment = data[:split]
                matches = find_tags(segment)
                if matches:
                    formula_count += matches.count(b"f")
                    merged_matches = matches.count(b"mergeCell")
                    drawing_matches = matches.count(b"drawing")
                    merged_range_count += merged_matches
                    hyperlink_count += matches.count(b"hyperlink")
                else:
                    merged_matches = 0
                    drawing_matches = 0
                if merged_matches or drawing_matches:
                    _collect_metadata_tags(
                        segment,
                        merged_ranges,
                        drawing_relationship_ids,
                        max_merged_ranges,
                        path,
                    )
                tail = data[split:]
    except KeyError as error:
        raise _invalid_workbook(path) from error
    matches = find_tags(tail)
    formula_count += matches.count(b"f")
    merged_matches = matches.count(b"mergeCell")
    drawing_matches = matches.count(b"drawing")
    merged_range_count += merged_matches
    hyperlink_count += matches.count(b"hyperlink")
    if merged_matches or drawing_matches:
        _collect_metadata_tags(
            tail,
            merged_ranges,
            drawing_relationship_ids,
            max_merged_ranges,
            path,
        )
    return (
        {
            "formulaCount": formula_count,
            "mergedRangeCount": merged_range_count,
            "hyperlinkCount": hyperlink_count,
        },
        merged_ranges,
        drawing_relationship_ids,
    )


def _cell_payload(cell: CellData, include_fields: list[str]) -> dict[str, object]:
    payload: dict[str, object] = {"address": cell.address}
    if "values" in include_fields:
        payload["value"] = cell.value
    if "formulas" in include_fields and cell.formula is not None:
        payload["formula"] = cell.formula
    if "cachedValues" in include_fields:
        payload["cachedValue"] = cell.cached_value
    if "numberFormats" in include_fields:
        payload["numberFormat"] = cell.number_format
    if "dataTypes" in include_fields:
        payload["dataType"] = cell.data_type
    if "sourceValues" in include_fields:
        payload["sourceValue"] = cell.source_value
    return payload


def _resolve_worksheets(archive: ZipFile, path: Path) -> list[WorksheetSource]:
    root = _read_xml(archive, "xl/workbook.xml")
    relationships = _read_workbook_relationships(archive)
    worksheets = []
    for position, sheet in enumerate(root.findall(f".//{{{MAIN_NS}}}sheet"), start=1):
        name = sheet.attrib.get("name")
        relationship_id = sheet.attrib.get(f"{{{RELATIONSHIP_NS}}}id")
        if name is None or relationship_id is None:
            raise _invalid_workbook(path)
        target = relationships.get(relationship_id)
        if target is None:
            raise _invalid_workbook(path)
        worksheets.append(
            WorksheetSource(
                name=name,
                path=target,
                hidden=sheet.attrib.get("state", "visible") != "visible",
                state=sheet.attrib.get("state", "visible"),
                position=position,
            )
        )
    return worksheets


def _read_workbook_relationships(archive: ZipFile) -> dict[str, str]:
    return _read_relationships(archive, "xl/_rels/workbook.xml.rels")


def _read_relationships(
    archive: ZipFile,
    relationships_path: str,
    *,
    missing_ok: bool = False,
) -> dict[str, str]:
    try:
        root = _read_xml(archive, relationships_path)
    except KeyError:
        if missing_ok:
            return {}
        raise
    relationships: dict[str, str] = {}
    for relationship in root.findall(f"{{{PACKAGE_RELATIONSHIP_NS}}}Relationship"):
        relationship_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        if relationship_id is None or target is None:
            continue
        relationships[relationship_id] = _normalize_relationship_target(
            relationships_path,
            target,
        )
    return relationships


def _normalize_relationship_target(relationships_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base_path = relationships_path.rsplit("/_rels/", 1)[0]
    return posixpath.normpath(posixpath.join(base_path, target))


def _relationships_path(source_path: str) -> str:
    parent, _, filename = source_path.rpartition("/")
    if parent:
        return f"{parent}/_rels/{filename}.rels"
    return f"_rels/{filename}.rels"


def _read_dimension(archive: ZipFile, worksheet_path: str) -> str:
    with archive.open(worksheet_path) as handle:
        chunk = handle.read(4096).decode("utf-8", errors="ignore")
    match = re.search(r'<dimension\s+ref="([^"]+)"', chunk)
    if match:
        try:
            bounds = range_boundaries(match.group(1))
            if all(value is not None and value > 0 for value in bounds):
                return match.group(1)
        except ValueError:
            pass

    # Worksheet dimensions are optional. Missing metadata must not hide cells
    # from agents that use inspection to choose their read range.
    min_row = min_column = None
    max_row = max_column = 0
    current_row = current_column = 0
    parser = xml.parsers.expat.ParserCreate(namespace_separator="}")

    def start_element(name: str, attributes: dict[str, str]) -> None:
        nonlocal min_row, min_column, max_row, max_column
        nonlocal current_row, current_column
        if name == f"{MAIN_NS}}}row":
            current_row = int(attributes.get("r", current_row + 1))
            current_column = 0
            return
        if name != f"{MAIN_NS}}}c":
            return
        address = attributes.get("r", "")
        if address:
            cell = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", address)
            if cell is None:
                raise ValueError(f"Invalid cell address {address!r}")
            column = column_index_from_string(cell.group(1))
            row = int(cell.group(2))
        else:
            column = current_column + 1
            row = current_row
        current_column = column
        min_row = row if min_row is None else min(min_row, row)
        min_column = column if min_column is None else min(min_column, column)
        max_row = max(max_row, row)
        max_column = max(max_column, column)

    parser.StartElementHandler = start_element
    try:
        with archive.open(worksheet_path) as handle:
            parser.ParseFile(handle)
    except (xml.parsers.expat.ExpatError, ValueError) as error:
        raise _invalid_workbook(Path(str(archive.filename))) from error
    if min_row is None or min_column is None:
        return "A1"
    first = f"{get_column_letter(min_column)}{min_row}"
    last = f"{get_column_letter(max_column)}{max_row}"
    return first if first == last else f"{first}:{last}"


def _dimension_size(dimension: str) -> tuple[int, int]:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(dimension)
    except ValueError:
        return 1, 1
    if None in {min_col, min_row, max_col, max_row}:
        return 1, 1
    return int(max_row), int(max_col)


class _SharedStringScanner:
    """Collect shared string values, including rich-text runs."""

    def __init__(
        self,
        parser: xml.parsers.expat.XMLParserType,
        strings_out: list[str],
    ) -> None:
        self._parser = parser
        self._strings_out = strings_out
        self._parts: list[str] | None = None
        self._capture: list[str] | None = None
        parser.StartElementHandler = self._start
        parser.EndElementHandler = self._end

    def _start(self, name: str, attrs: dict[str, str]) -> None:
        local = _LOCAL_NAMES.get(name)
        if local is None:
            local = name.rpartition(":")[2]
            _LOCAL_NAMES[name] = local
        if local == "t":
            if self._parts is not None:
                self._capture = []
                self._parser.CharacterDataHandler = self._capture.append
        elif local == "si":
            self._parts = []

    def _end(self, name: str) -> None:
        local = _LOCAL_NAMES.get(name)
        if local is None:
            local = name.rpartition(":")[2]
            _LOCAL_NAMES[name] = local
        if local == "t":
            capture = self._capture
            if capture is not None and self._parts is not None:
                self._parts.append("".join(capture))
                self._capture = None
                self._parser.CharacterDataHandler = None
        elif local == "si":
            if self._parts is not None:
                self._strings_out.append("".join(self._parts))
                self._parts = None


def _read_shared_strings(archive: ZipFile) -> list[str]:
    try:
        info = archive.getinfo("xl/sharedStrings.xml")
    except KeyError:
        return []

    strings: list[str] = []
    parser = xml.parsers.expat.ParserCreate()
    parser.buffer_text = True
    _SharedStringScanner(parser, strings)
    with archive.open(info) as handle:
        while True:
            chunk = handle.read(_CHUNK_SIZE)
            if not chunk:
                parser.Parse(b"", True)
                break
            parser.Parse(chunk, False)
    return strings


def _read_number_formats(archive: ZipFile) -> dict[int, str]:
    try:
        root = _read_xml(archive, "xl/styles.xml")
    except KeyError:
        return {}

    custom_formats: dict[int, str] = {}
    for num_fmt in root.findall(f".//{{{MAIN_NS}}}numFmt"):
        num_fmt_id = num_fmt.attrib.get("numFmtId")
        format_code = num_fmt.attrib.get("formatCode")
        if num_fmt_id is None or format_code is None:
            continue
        try:
            custom_formats[int(num_fmt_id)] = format_code
        except ValueError:
            continue

    cell_formats: dict[int, str] = {}
    cell_xfs = root.find(f".//{{{MAIN_NS}}}cellXfs")
    if cell_xfs is None:
        return cell_formats
    for style_index, xf in enumerate(cell_xfs):
        num_fmt_id = xf.attrib.get("numFmtId")
        if num_fmt_id is None:
            continue
        try:
            parsed_id = int(num_fmt_id)
        except ValueError:
            continue
        cell_formats[style_index] = custom_formats.get(
            parsed_id,
            BUILTIN_FORMATS.get(parsed_id, "General"),
        )
    return cell_formats


def _parse_cell_range(cell_range: str) -> tuple[int, int, int, int]:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    except ValueError as error:
        raise SheetOpsError(
            "invalid_range",
            "Cell range must be a bounded A1-style range",
            {"range": cell_range},
        ) from error
    if None in {min_col, min_row, max_col, max_row}:
        raise SheetOpsError(
            "invalid_range",
            "Cell range must be a bounded A1-style range",
            {"range": cell_range},
        )
    return int(min_col), int(min_row), int(max_col), int(max_row)


def _column_letter_to_index(column: str, field_name: str) -> int:
    try:
        return column_index_from_string(column)
    except ValueError as error:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "column letter"},
        ) from error


def _read_xml(archive: ZipFile, name: str) -> ET.Element:
    return ET.fromstring(archive.read(name))


def _local_name(tag: str) -> str:
    if "}" not in tag:
        return tag
    return tag.rsplit("}", 1)[1]


def _invalid_workbook(path: Path) -> SheetOpsError:
    return SheetOpsError(
        "invalid_workbook",
        "Workbook could not be opened",
        {"path": str(path)},
    )
