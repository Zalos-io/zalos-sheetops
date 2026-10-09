import posixpath
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.cellref import coordinate_to_tuple, range_boundaries

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class WorksheetSource:
    def __init__(self, name: str, path: str) -> None:
        self.name = name
        self.path = path


class HyperlinkInfo:
    def __init__(
        self,
        target: str | None,
        display: str | None,
        tooltip: str | None,
    ) -> None:
        self.target = target
        self.display = display
        self.tooltip = tooltip


def assert_hyperlinks(
    path: Path,
    sheet_name: str,
    checks: list[dict[str, object]],
) -> dict[str, object]:
    validated_checks = _validate_checks(checks)
    target_cells = {_normalized_cell(str(check["cell"])) for check in validated_checks}
    try:
        with ZipFile(path, "r") as archive:
            worksheet = _resolve_worksheet(archive, sheet_name, path)
            shared_strings = _read_shared_strings(archive)
            hyperlinks = _read_target_hyperlinks(
                archive,
                worksheet.path,
                target_cells,
                shared_strings,
                path,
            )
    except BadZipFile as error:
        raise _invalid_workbook(path) from error
    except ET.ParseError as error:
        raise _invalid_workbook(path) from error
    except KeyError as error:
        raise _invalid_workbook(path) from error

    failures: list[dict[str, object]] = []
    for check in validated_checks:
        cell = _normalized_cell(str(check["cell"]))
        hyperlink = hyperlinks.get(cell)
        if hyperlink is None:
            failures.append(_missing_hyperlink_failure(cell, check))
            continue
        failures.extend(_check_field(cell, "target", check, hyperlink.target))
        failures.extend(_check_field(cell, "display", check, hyperlink.display))
        failures.extend(_check_field(cell, "tooltip", check, hyperlink.tooltip))

    return {
        "valid": not failures,
        "sheet": sheet_name,
        "checkedCount": len(validated_checks),
        "failureCount": len(failures),
        "failures": failures,
    }


def _validate_checks(checks: list[dict[str, object]]) -> list[dict[str, object]]:
    if not isinstance(checks, list) or not checks:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "checks", "expected": "non-empty list[hyperlinkCheck]"},
        )
    for index, check in enumerate(checks):
        if not isinstance(check, dict):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": f"checks[{index}]", "expected": "hyperlinkCheck"},
            )
        cell = check.get("cell")
        if not isinstance(cell, str):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {"field": f"checks[{index}].cell", "expected": "A1-style cell"},
            )
        _normalized_cell(cell)
        expected_fields = ["target", "display", "tooltip"]
        if not any(field in check for field in expected_fields):
            raise SheetOpsError(
                "invalid_field",
                "Field has invalid type",
                {
                    "field": f"checks[{index}]",
                    "expected": (
                        "hyperlinkCheck with cell and at least one of "
                        "target, display, tooltip"
                    ),
                },
            )
        for field in expected_fields:
            value = check.get(field)
            if value is not None and not isinstance(value, str):
                raise SheetOpsError(
                    "invalid_field",
                    "Field has invalid type",
                    {
                        "field": f"checks[{index}].{field}",
                        "expected": "string | null",
                    },
                )
    return checks


def _missing_hyperlink_failure(
    cell: str,
    check: dict[str, object],
) -> dict[str, object]:
    expected = {
        field: check[field]
        for field in ("target", "display", "tooltip")
        if field in check
    }
    expected_value: object = (
        next(iter(expected.values())) if len(expected) == 1 else expected
    )
    return {
        "cell": cell,
        "code": "missing_hyperlink",
        "expected": expected_value,
        "actual": None,
    }


def _check_field(
    cell: str,
    field: str,
    check: dict[str, object],
    actual: str | None,
) -> list[dict[str, object]]:
    expected = check.get(field)
    if not isinstance(expected, str) or expected == actual:
        return []
    return [
        {
            "cell": cell,
            "code": f"{field}_mismatch",
            "expected": expected,
            "actual": actual,
        }
    ]


def _resolve_worksheet(
    archive: ZipFile,
    sheet_name: str,
    path: Path,
) -> WorksheetSource:
    workbook_root = _read_xml(archive, "xl/workbook.xml")
    relationships = _read_relationships(archive, "xl/_rels/workbook.xml.rels", "xl")
    for sheet in workbook_root.findall(f".//{{{MAIN_NS}}}sheet"):
        name = sheet.attrib.get("name")
        relationship_id = sheet.attrib.get(f"{{{RELATIONSHIP_NS}}}id")
        if name != sheet_name:
            continue
        if relationship_id is None or relationship_id not in relationships:
            raise _invalid_workbook(path)
        return WorksheetSource(name, relationships[relationship_id])
    raise SheetOpsError(
        "invalid_sheet",
        "Requested sheet does not exist",
        {"sheet": sheet_name},
    )


def _read_target_hyperlinks(
    archive: ZipFile,
    worksheet_path: str,
    target_cells: set[str],
    shared_strings: list[str],
    path: Path,
) -> dict[str, HyperlinkInfo]:
    relationship_targets = _read_worksheet_relationships(archive, worksheet_path)
    cell_values: dict[str, str] = {}
    hyperlinks: dict[str, HyperlinkInfo] = {}
    try:
        with archive.open(worksheet_path) as handle:
            for _, element in ET.iterparse(handle, events=("end",)):
                local_name = _local_name(element.tag)
                if local_name == "c":
                    coordinate = element.attrib.get("r")
                    if coordinate is not None:
                        cell = _normalized_cell(coordinate)
                        if cell in target_cells:
                            cell_values[cell] = _cell_text(element, shared_strings)
                    element.clear()
                    continue
                if local_name != "hyperlink":
                    continue
                ref = element.attrib.get("ref")
                if ref is None:
                    element.clear()
                    continue
                matching_cells = [
                    cell
                    for cell in sorted(target_cells)
                    if _cell_in_reference(cell, ref)
                ]
                if not matching_cells:
                    element.clear()
                    continue
                target = _hyperlink_target(element, relationship_targets)
                tooltip = element.attrib.get("tooltip")
                display = element.attrib.get("display")
                for cell in matching_cells:
                    hyperlinks[cell] = HyperlinkInfo(
                        target,
                        display or cell_values.get(cell),
                        tooltip,
                    )
                element.clear()
    except KeyError as error:
        raise _invalid_workbook(path) from error
    return hyperlinks


def _read_worksheet_relationships(
    archive: ZipFile,
    worksheet_path: str,
) -> dict[str, str]:
    rels_path = _worksheet_relationship_path(worksheet_path)
    try:
        return _read_relationships(
            archive,
            rels_path,
            posixpath.dirname(worksheet_path),
        )
    except KeyError:
        return {}


def _hyperlink_target(
    element: ET.Element,
    relationship_targets: dict[str, str],
) -> str | None:
    location = element.attrib.get("location")
    if location is not None:
        return location if location.startswith("#") else f"#{location}"
    relationship_id = element.attrib.get(f"{{{RELATIONSHIP_NS}}}id")
    if relationship_id is None:
        return None
    return relationship_targets.get(relationship_id)


def _cell_text(element: ET.Element, shared_strings: list[str]) -> str:
    cell_type = element.attrib.get("t")
    if cell_type == "inlineStr":
        inline_string = _find_child(element, "is")
        return _text_content(inline_string) if inline_string is not None else ""
    value = _find_child(element, "v")
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return ""
    return value.text


def _read_shared_strings(archive: ZipFile) -> list[str]:
    try:
        root = _read_xml(archive, "xl/sharedStrings.xml")
    except KeyError:
        return []
    strings: list[str] = []
    for item in root:
        if _local_name(item.tag) != "si":
            continue
        strings.append(_text_content(item))
    return strings


def _read_relationships(
    archive: ZipFile,
    rels_path: str,
    base_path: str,
) -> dict[str, str]:
    relationships: dict[str, str] = {}
    for relationship in _read_relationship_elements(archive, rels_path):
        relationship_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        if relationship_id is None or target is None:
            continue
        relationships[relationship_id] = _normalize_relationship_target(
            target,
            base_path,
        )
    return relationships


def _read_relationship_elements(
    archive: ZipFile,
    rels_path: str,
) -> list[ET.Element]:
    root = _read_xml(archive, rels_path)
    return [
        relationship
        for relationship in root
        if _local_name(relationship.tag) == "Relationship"
    ]


def _read_xml(archive: ZipFile, name: str) -> ET.Element:
    with archive.open(name) as handle:
        return ET.parse(handle).getroot()


def _worksheet_relationship_path(worksheet_path: str) -> str:
    directory = posixpath.dirname(worksheet_path)
    filename = posixpath.basename(worksheet_path)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _normalize_relationship_target(target: str, base_path: str) -> str:
    if "://" in target or target.startswith("#") or target.startswith("mailto:"):
        return target
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_path, target))


def _cell_in_reference(cell: str, reference: str) -> bool:
    if ":" not in reference:
        return _normalized_cell(reference) == cell
    min_col, min_row, max_col, max_row = range_boundaries(reference)
    row, column = coordinate_to_tuple(cell)
    return min_row <= row <= max_row and min_col <= column <= max_col


def _normalized_cell(cell: str) -> str:
    try:
        row, column = coordinate_to_tuple(cell)
    except ValueError as error:
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": "cell", "expected": "A1-style cell"},
        ) from error
    letters = ""
    value = column
    while value:
        value, remainder = divmod(value - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def _find_child(element: ET.Element, local_name: str) -> ET.Element | None:
    for child in element:
        if _local_name(child.tag) == local_name:
            return child
    return None


def _text_content(element: ET.Element) -> str:
    return "".join(text for text in element.itertext())


def _local_name(tag: str) -> str:
    if "}" not in tag:
        return tag
    return tag.rsplit("}", 1)[1]


def _invalid_workbook(path: Path) -> SheetOpsError:
    return SheetOpsError(
        "invalid_workbook",
        "Workbook file could not be read",
        {"path": str(path)},
    )
