import posixpath
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.cellref import coordinate_to_tuple, get_column_letter

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class WorksheetSource:
    def __init__(self, name: str, path: str) -> None:
        self.name = name
        self.path = path


def assert_styles_xml(
    path: Path,
    sheet_name: str,
    checks: list[dict[str, object]],
) -> dict[str, object]:
    try:
        with ZipFile(path, "r") as archive:
            worksheet = _resolve_worksheet(archive, sheet_name, path)
            target_cells = {_normalized_cell(str(check["cell"])) for check in checks}
            style_fills = _read_style_fills(archive)
            cells = _read_target_cells(
                archive,
                worksheet.path,
                target_cells,
                style_fills,
                path,
            )
            comments = _read_comments_for_worksheet(archive, worksheet.path)
    except BadZipFile as error:
        raise _invalid_workbook(path) from error
    except ET.ParseError as error:
        raise _invalid_workbook(path) from error
    except KeyError as error:
        raise _invalid_workbook(path) from error

    failures = []
    for check in checks:
        cell = _normalized_cell(str(check["cell"]))
        actual = cells.get(cell, {"formula": False, "fill": None})
        failures.extend(_check_formula(cell, check, bool(actual["formula"])))
        failures.extend(_check_fill(cell, check, actual["fill"]))
        failures.extend(_check_comment(cell, check, comments.get(cell)))

    return {
        "valid": not failures,
        "sheet": sheet_name,
        "checkedCount": len(checks),
        "failureCount": len(failures),
        "failures": failures,
    }


def can_assert_styles_xml(checks: list[dict[str, object]]) -> bool:
    supported = {"cell", "fill", "commentContains", "formula"}
    return all(set(check).issubset(supported) for check in checks)


def _check_formula(
    cell: str,
    check: dict[str, object],
    actual_formula: bool,
) -> list[dict[str, object]]:
    expected_formula = check.get("formula")
    if not isinstance(expected_formula, bool) or actual_formula == expected_formula:
        return []
    return [
        {
            "cell": cell,
            "code": "formula_mismatch",
            "expected": expected_formula,
            "actual": actual_formula,
        }
    ]


def _check_fill(
    cell: str,
    check: dict[str, object],
    actual_fill: object,
) -> list[dict[str, object]]:
    fill = check.get("fill")
    if not isinstance(fill, str):
        return []
    expected_fill = _normalize_color(fill, "fill")
    if actual_fill == expected_fill:
        return []
    return [
        {
            "cell": cell,
            "code": "fill_mismatch",
            "expected": expected_fill,
            "actual": actual_fill,
        }
    ]


def _check_comment(
    cell: str,
    check: dict[str, object],
    actual_comment: str | None,
) -> list[dict[str, object]]:
    comment_contains = check.get("commentContains")
    if not isinstance(comment_contains, str):
        return []
    if actual_comment is not None and comment_contains in actual_comment:
        return []
    return [
        {
            "cell": cell,
            "code": "comment_mismatch",
            "expected": comment_contains,
            "actual": actual_comment,
        }
    ]


def _resolve_worksheet(
    archive: ZipFile,
    sheet_name: str,
    path: Path,
) -> WorksheetSource:
    workbook_root = _read_xml(archive, "xl/workbook.xml")
    relationships = _read_relationships(archive, "xl/_rels/workbook.xml.rels")
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


def _read_style_fills(archive: ZipFile) -> list[str | None]:
    try:
        root = _read_xml(archive, "xl/styles.xml")
    except KeyError:
        return []
    fills = _read_fills(root)
    style_fills: list[str | None] = []
    cell_xfs = _find_child(root, "cellXfs")
    if cell_xfs is None:
        return style_fills
    for style in cell_xfs:
        if _local_name(style.tag) != "xf":
            continue
        fill_id = _int_attribute(style, "fillId")
        if fill_id is None or fill_id < 0 or fill_id >= len(fills):
            style_fills.append(None)
            continue
        style_fills.append(fills[fill_id])
    return style_fills


def _read_fills(root: ET.Element) -> list[str | None]:
    fills_element = _find_child(root, "fills")
    if fills_element is None:
        return []
    fills: list[str | None] = []
    for fill in fills_element:
        if _local_name(fill.tag) != "fill":
            continue
        pattern_fill = _find_child(fill, "patternFill")
        if pattern_fill is None or pattern_fill.attrib.get("patternType") is None:
            fills.append(None)
            continue
        fg_color = _find_child(pattern_fill, "fgColor")
        fills.append(_rgb_color(fg_color))
    return fills


def _read_target_cells(
    archive: ZipFile,
    worksheet_path: str,
    target_cells: set[str],
    style_fills: list[str | None],
    path: Path,
) -> dict[str, dict[str, object]]:
    cells: dict[str, dict[str, object]] = {}
    try:
        with archive.open(worksheet_path) as handle:
            for _, cell in ET.iterparse(handle, events=("end",)):
                if _local_name(cell.tag) != "c":
                    continue
                coordinate = cell.attrib.get("r")
                if coordinate is None:
                    cell.clear()
                    continue
                coordinate = _normalized_cell(coordinate)
                if coordinate not in target_cells:
                    cell.clear()
                    continue
                style_id = _int_attribute(cell, "s")
                cells[coordinate] = {
                    "formula": _find_child(cell, "f") is not None,
                    "fill": _fill_for_style_id(style_fills, style_id),
                }
                cell.clear()
    except KeyError as error:
        raise _invalid_workbook(path) from error
    return cells


def _read_comments_for_worksheet(
    archive: ZipFile,
    worksheet_path: str,
) -> dict[str, str]:
    comments_path = _comments_path_for_worksheet(archive, worksheet_path)
    if comments_path is None:
        return {}
    try:
        root = _read_xml(archive, comments_path)
    except KeyError:
        return {}
    comments: dict[str, str] = {}
    for comment in root.iter():
        if _local_name(comment.tag) != "comment":
            continue
        ref = comment.attrib.get("ref")
        if ref is None:
            continue
        comments[_normalized_cell(ref)] = _text_content(comment)
    return comments


def _comments_path_for_worksheet(
    archive: ZipFile,
    worksheet_path: str,
) -> str | None:
    rels_path = _worksheet_relationship_path(worksheet_path)
    try:
        _read_xml(archive, rels_path)
    except KeyError:
        return None
    for relationship in _read_relationship_elements(archive, rels_path):
        relationship_type = relationship.attrib.get("Type", "")
        if not relationship_type.endswith("/comments"):
            continue
        target = relationship.attrib.get("Target")
        if target is None:
            continue
        return _normalize_relationship_target(target, posixpath.dirname(worksheet_path))
    return None


def _read_relationships(archive: ZipFile, rels_path: str) -> dict[str, str]:
    relationships: dict[str, str] = {}
    for relationship in _read_relationship_elements(archive, rels_path):
        relationship_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        if relationship_id is None or target is None:
            continue
        relationships[relationship_id] = _normalize_relationship_target(target, "xl")
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


def _worksheet_relationship_path(worksheet_path: str) -> str:
    directory = posixpath.dirname(worksheet_path)
    filename = posixpath.basename(worksheet_path)
    return posixpath.join(directory, "_rels", f"{filename}.rels")


def _normalize_relationship_target(target: str, base_path: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_path, target))


def _fill_for_style_id(
    style_fills: list[str | None],
    style_id: int | None,
) -> str | None:
    if style_id is None or style_id < 0 or style_id >= len(style_fills):
        return None
    return style_fills[style_id]


def _rgb_color(color: ET.Element | None) -> str | None:
    if color is None:
        return None
    rgb = color.attrib.get("rgb")
    if rgb is None:
        return None
    return _normalize_color(rgb, "fill")


def _normalize_color(value: str, field_name: str) -> str:
    color = value.strip().upper()
    if color.startswith("#"):
        color = color[1:]
    if len(color) == 8:
        color = color[2:]
    if len(color) != 6 or any(char not in "0123456789ABCDEF" for char in color):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {"field": field_name, "expected": "hex color string"},
        )
    return f"#{color}"


def _text_content(element: ET.Element) -> str:
    parts: list[str] = []
    for child in element.iter():
        if _local_name(child.tag) == "t" and child.text is not None:
            parts.append(child.text)
    return "".join(parts)


def _int_attribute(element: ET.Element, name: str) -> int | None:
    value = element.attrib.get(name)
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _normalized_cell(cell: str) -> str:
    row, column = coordinate_to_tuple(cell)
    return f"{get_column_letter(column)}{row}"


def _read_xml(archive: ZipFile, name: str) -> ET.Element:
    return ET.fromstring(archive.read(name))


def _find_child(element: ET.Element, local_name: str) -> ET.Element | None:
    for child in element:
        if _local_name(child.tag) == local_name:
            return child
    return None


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
