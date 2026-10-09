import posixpath
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from zalos_sheetops.errors import SheetOpsError

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass(frozen=True)
class WorksheetMetadata:
    name: str
    path: str
    hidden: bool


def read_worksheet_metadata(path: Path) -> list[WorksheetMetadata]:
    try:
        with ZipFile(path, "r") as archive:
            root = _read_xml(archive, "xl/workbook.xml")
            relationships = _read_workbook_relationships(archive)
            worksheets: list[WorksheetMetadata] = []
            for sheet in root.findall(f".//{{{MAIN_NS}}}sheet"):
                name = sheet.attrib.get("name")
                relationship_id = sheet.attrib.get(f"{{{RELATIONSHIP_NS}}}id")
                if name is None or relationship_id is None:
                    raise _invalid_workbook(path)
                target = relationships.get(relationship_id)
                if target is None:
                    raise _invalid_workbook(path)
                worksheets.append(
                    WorksheetMetadata(
                        name=name,
                        path=target,
                        hidden=sheet.attrib.get("state", "visible") != "visible",
                    )
                )
            return worksheets
    except (BadZipFile, ET.ParseError, KeyError) as error:
        raise _invalid_workbook(path) from error


def read_column_widths(path: Path, worksheet_path: str) -> dict[int, float]:
    try:
        with ZipFile(path, "r") as archive:
            widths: dict[int, float] = {}
            with archive.open(worksheet_path) as handle:
                for event, element in ET.iterparse(handle, events=("start", "end")):
                    if event == "end" and _local_name(element.tag) == "cols":
                        return widths
                    if event != "start" or _local_name(element.tag) != "col":
                        continue
                    minimum = _positive_int_attribute(element, "min")
                    maximum = _positive_int_attribute(element, "max")
                    width = _float_attribute(element, "width")
                    if minimum is None or maximum is None or width is None:
                        continue
                    for column_index in range(minimum, maximum + 1):
                        widths[column_index] = width
            return widths
    except (BadZipFile, ET.ParseError, KeyError) as error:
        raise _invalid_workbook(path) from error


def _read_workbook_relationships(archive: ZipFile) -> dict[str, str]:
    root = _read_xml(archive, "xl/_rels/workbook.xml.rels")
    relationships: dict[str, str] = {}
    for relationship in root:
        if _local_name(relationship.tag) != "Relationship":
            continue
        relationship_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        if relationship_id is None or target is None:
            continue
        relationships[relationship_id] = _normalize_relationship_target(target)
    return relationships


def _normalize_relationship_target(target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def _read_xml(archive: ZipFile, name: str) -> ET.Element:
    with archive.open(name) as handle:
        return ET.parse(handle).getroot()


def _local_name(tag: str) -> str:
    if "}" not in tag:
        return tag
    return tag.rsplit("}", 1)[1]


def _positive_int_attribute(element: ET.Element, name: str) -> int | None:
    value = element.attrib.get(name)
    if value is None:
        return None
    try:
        parsed = int(value)
    except ValueError:
        return None
    if parsed < 1:
        return None
    return parsed


def _float_attribute(element: ET.Element, name: str) -> float | None:
    value = element.attrib.get(name)
    if value is None:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _invalid_workbook(path: Path) -> SheetOpsError:
    return SheetOpsError(
        "invalid_workbook",
        "Workbook could not be opened",
        {"path": str(path)},
    )
