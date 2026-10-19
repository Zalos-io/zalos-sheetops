"""Reviewed style requirements for the fixed public task sample, with visible colors."""

import colorsys
from copy import copy
from decimal import Decimal, localcontext
from xml.etree import ElementTree as ET

from openpyxl.cell.cell import Cell
from openpyxl.styles.colors import COLOR_INDEX, Color

from .scoring import canonical

# A mention of existing formatting or bold input is not a requested style change.
STYLE_REQUIREMENTS = {
    "51680": [
        "fontName",
        "fontSize",
        "bold",
        "italic",
        "fill",
        "numberFormat",
        "alignment",
        "border",
    ],
    "192-22": [],
    "47827": ["bold", "fill"],
    "50682": [],
    "343-20": ["alignment"],
    "54638": ["fontName", "fill", "border"],
    "120-24": ["fill"],
    "510-3": ["numberFormat", "bold"],
    "398-14": ["fontName", "fontSize", "bold", "fontColor", "alignment"],
}

ORACLE_ISSUES = {
    "402-43": {
        "status": "ambiguous_task_metadata",
        "reason": "Answer metadata requires new Sheet2, "
        "which the task instruction does not name",
    },
    "395-36": {
        "status": "ambiguous_task_metadata",
        "reason": "Answer metadata requires new MyResult, while the instruction "
        "asks to modify Result what I am getting",
    },
    "168-17": {
        "status": "oracle_preservation_conflict",
        "reason": "Golden workbook deletes unrelated Final result sheet; "
        "the instruction only asks to delete preceding rows in Statement",
    },
}

# Reviewed after the 2026-10 cohorts. Graded as unscorable, but kept out of
# ORACLE_ISSUES so seeded selection still reproduces the published 33 tasks.
REFERENCE_CONFLICTS = {
    "54638": {
        "status": "oracle_style_conflict",
        "reason": "Golden keeps font 等线 on every cell, while the instruction asks "
        "to recreate Sheet1 with Calibri font; fontName enforces the golden font",
    },
    "398-14": {
        "status": "oracle_style_conflict",
        "reason": "Golden sets font 等线 on every cell of all five sheets, while the "
        "instruction asks for font Calibri 11pts; fontName enforces the golden font",
    },
    "48975": {
        "status": "oracle_preservation_conflict",
        "reason": "Golden filters Output column B only and keeps A11:A17 as "
        "unfiltered Input links; the instruction asks to copy columns A and B",
    },
}


def reviewed_task(task: dict) -> dict:
    key = str(task["id"])
    result = task.copy()
    if key in STYLE_REQUIREMENTS:
        result.update(
            styleComponents=STYLE_REQUIREMENTS[key],
            styleRubric="reviewed-public-sample-v1",
        )
    issue = ORACLE_ISSUES.get(key) or REFERENCE_CONFLICTS.get(key)
    if issue:
        result["oracleIssue"] = issue.copy()
    return result


def values_equal(expected: object, actual: object) -> bool:
    if expected in (None, "") and actual in (None, ""):
        return True
    if isinstance(expected, list) and isinstance(actual, list):
        return len(expected) == len(actual) and all(
            values_equal(a, b) for a, b in zip(expected, actual)
        )
    if type(expected) in {int, float} and type(actual) in {int, float}:
        if isinstance(expected, float) or isinstance(actual, float):
            # Excel uses 15 significant decimal digits; retain exact integer IDs.
            # https://support.microsoft.com/en-us/excel/change-formula-recalculation-iteration-or-precision-in-excel
            with localcontext() as context:
                context.prec = 15
                return +Decimal(str(expected)) == +Decimal(str(actual))
    return canonical(expected) == canonical(actual)


def visible_color(color: Color | None, cell: Cell) -> object:
    if color is None:
        return None
    if color.type == "rgb":
        return color.rgb[-6:].upper()
    if color.type == "indexed":
        if 0 <= color.indexed < len(COLOR_INDEX):
            return COLOR_INDEX[color.indexed][-6:].upper()
        return ("indexed", color.indexed)
    if color.type != "theme" or not cell.parent.parent.loaded_theme:
        return (color.type, color.value, color.tint)
    theme = ET.fromstring(cell.parent.parent.loaded_theme)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = theme.find("a:themeElements/a:clrScheme", ns)
    slots = [
        "lt1",
        "dk1",
        "lt2",
        "dk2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
    ]
    if scheme is None or not 0 <= color.theme < len(slots):
        return (color.type, color.value, color.tint)
    slot = scheme.find("a:" + slots[color.theme], ns)
    if slot is None or not len(slot):
        return (color.type, color.value, color.tint)
    source = list(slot)[0]
    rgb = source.get("lastClr") or source.get("val")
    values = [int(rgb[n : n + 2], 16) / 255 for n in (0, 2, 4)]
    hue, light, saturation = colorsys.rgb_to_hls(*values)
    tint = float(color.tint or 0)
    light = light * (1 + tint) if tint < 0 else light * (1 - tint) + tint
    tinted = colorsys.hls_to_rgb(hue, light, saturation)
    return "".join(f"{round(channel * 255):02X}" for channel in tinted)


def style_component(cell: Cell, component: str) -> object:
    if component == "fill":
        fill = cell.fill
        return (
            (fill.patternType, visible_color(fill.fgColor, cell))
            if fill.patternType
            else None
        )
    if component == "border":
        border = cell.border
        sides = []
        for name in [
            "left",
            "right",
            "top",
            "bottom",
            "diagonal",
            "vertical",
            "horizontal",
            "start",
            "end",
        ]:
            side = getattr(border, name)
            sides.append(
                (side.style, visible_color(side.color, cell))
                if side and side.style
                else None
            )
        return (tuple(sides), bool(border.diagonalUp), bool(border.diagonalDown))
    getters = {
        "fontName": lambda: cell.font.name,
        "fontSize": lambda: cell.font.sz,
        "bold": lambda: bool(cell.font.b),
        "italic": lambda: bool(cell.font.i),
        "fontColor": lambda: visible_color(cell.font.color, cell),
        "numberFormat": lambda: cell.number_format,
        "alignment": lambda: copy(cell.alignment),
    }
    if component not in getters:
        raise ValueError(f"Unknown style component: {component}")
    return getters[component]()


def style_matches(expected: Cell, actual: Cell, components: list[str]) -> bool:
    return all(
        style_component(expected, part) == style_component(actual, part)
        for part in components
    )
