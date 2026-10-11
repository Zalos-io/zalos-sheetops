"""Recognize workbook payloads before a transport releases their bytes."""

from pathlib import PurePath
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile, is_zipfile

_METADATA_LIMIT = 256 * 1024
_WORKBOOK_TYPES = {
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "application/vnd.ms-excel.template.macroEnabled.main+xml",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
}


def classify_delivery(artifact_path: str, probe: object, **kwargs: object) -> dict:
    if PurePath(artifact_path).suffix.lower() in {
        ".xlsx",
        ".xlsm",
        ".xltx",
        ".xltm",
        ".xlsb",
    }:
        return {"action": "require_receipt"}
    try:
        if not is_zipfile(probe):
            return {"action": "allow"}
        with ZipFile(probe) as archive:
            names = archive.namelist()
            if any(name in names for name in ("xl/workbook.xml", "xl/workbook.bin")):
                return {"action": "require_receipt"}
            if "[Content_Types].xml" not in names:
                return {"action": "allow"}
            entries = [
                entry
                for entry in archive.infolist()
                if entry.filename == "[Content_Types].xml"
            ]
            if len(entries) != 1 or entries[0].file_size > _METADATA_LIMIT:
                raise ValueError("Archive metadata exceeds the delivery probe limit")
            with archive.open(entries[0]) as source:
                data = source.read(_METADATA_LIMIT + 1)
            if len(data) > _METADATA_LIMIT:
                raise ValueError("Archive metadata exceeds the delivery probe limit")
            root = ElementTree.fromstring(data)
            protected = any(node.get("ContentType") in _WORKBOOK_TYPES for node in root)
            return {"action": "require_receipt" if protected else "allow"}
    except (
        BadZipFile,
        ValueError,
        OSError,
        RuntimeError,
        ElementTree.ParseError,
    ) as exc:
        return {
            "action": "block",
            "message": f"Delivery content classification failed: {exc}",
        }
