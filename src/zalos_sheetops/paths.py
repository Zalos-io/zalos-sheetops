from pathlib import Path

from zalos_sheetops.errors import SheetOpsError


def validate_agent_path(raw_path: str) -> Path:
    if _contains_disallowed_path_chars(raw_path) or ".." in Path(raw_path).parts:
        raise SheetOpsError(
            "invalid_path",
            "Path contains disallowed characters or traversal",
            {"path": raw_path},
        )

    path = Path(raw_path).expanduser()
    if not path.exists():
        raise SheetOpsError(
            "file_not_found",
            "Workbook file does not exist",
            {"path": raw_path},
        )

    return path


def _contains_disallowed_path_chars(raw_path: str) -> bool:
    if any(char in raw_path for char in ("?", "#", "%")):
        return True
    return any(ord(char) < 0x20 for char in raw_path)
