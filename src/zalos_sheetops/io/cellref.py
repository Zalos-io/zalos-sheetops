"""Cell reference utilities for the XML engine.

Local implementations matching the openpyxl helpers previously used here, so
XML-engine commands do not have to import openpyxl (which drags in numpy and
dominates CLI startup time).
"""

MAX_COLUMN_INDEX = 18278  # "ZZZ"

# OOXML built-in number formats (ECMA-376 part 1, 18.8.30).
BUILTIN_FORMATS: dict[int, str] = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: '"$"#,##0_);("$"#,##0)',
    6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)',
    8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "mm-dd-yy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yy h:mm",
    37: "#,##0_);(#,##0)",
    38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    44: '_("$"* #,##0.00_)_("$"* \\(#,##0.00\\)_("$"* "-"??_)_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}


def coordinate_to_tuple(coordinate: str) -> tuple[int, int]:
    """Convert an A1-style coordinate to a (row, column) tuple."""
    column = 0
    for index, character in enumerate(coordinate):
        code = ord(character)
        if 65 <= code <= 90:
            column = column * 26 + code - 64
        elif 97 <= code <= 122:
            column = column * 26 + code - 96
        elif 48 <= code <= 57:
            if not column:
                break
            try:
                return int(coordinate[index:]), column
            except ValueError:
                break
        else:
            break
    raise ValueError(f"Invalid coordinate {coordinate}")


def column_index_from_string(column: str) -> int:
    """Convert an uppercase column letter (A..ZZZ) to a 1-based index."""
    index = 0
    for character in column:
        code = ord(character)
        if not 65 <= code <= 90:
            raise ValueError(f"{column} is not a valid column name")
        index = index * 26 + code - 64
    if not 1 <= index <= MAX_COLUMN_INDEX:
        raise ValueError(f"{column} is not a valid column name")
    return index


def get_column_letter(column_index: int) -> str:
    """Convert a 1-based column index to a column letter."""
    if not 1 <= column_index <= MAX_COLUMN_INDEX:
        raise ValueError(f"Invalid column index {column_index}")
    letters: list[str] = []
    while column_index:
        column_index, remainder = divmod(column_index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def range_boundaries(
    range_string: str,
) -> tuple[int | None, int | None, int | None, int | None]:
    """Return (min_col, min_row, max_col, max_row) for an A1-style range.

    Unbounded parts are returned as None, matching openpyxl: "A:C" yields
    (1, None, 3, None) and "1:3" yields (None, 1, None, 3). Invalid input
    raises ValueError.
    """
    if ":" in range_string:
        left, _, right = range_string.partition(":")
    else:
        left = right = range_string
    min_col, min_row = _parse_range_part(left, range_string)
    max_col, max_row = _parse_range_part(right, range_string)
    if (min_col is None) != (max_col is None) or (min_row is None) != (max_row is None):
        raise ValueError(f"{range_string} is not a valid coordinate or range")
    return min_col, min_row, max_col, max_row


def _parse_range_part(
    part: str,
    range_string: str,
) -> tuple[int | None, int | None]:
    letters: list[str] = []
    digits: list[str] = []
    for character in part:
        code = ord(character)
        if 97 <= code <= 122:
            character = character.upper()
            code -= 32
        if 65 <= code <= 90:
            if digits:
                raise ValueError(f"{range_string} is not a valid coordinate or range")
            letters.append(character)
        elif 48 <= code <= 57:
            digits.append(character)
        elif character != "$":
            raise ValueError(f"{range_string} is not a valid coordinate or range")
    if not letters and not digits:
        raise ValueError(f"{range_string} is not a valid coordinate or range")
    column = column_index_from_string("".join(letters)) if letters else None
    row = int("".join(digits)) if digits else None
    if row is not None and row < 1:
        raise ValueError(f"{range_string} is not a valid coordinate or range")
    return column, row
