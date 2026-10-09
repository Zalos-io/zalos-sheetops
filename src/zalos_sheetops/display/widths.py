# Width values and formulas are adapted from XlsxWriter's BSD-2-Clause
# utility.py/worksheet.py autofit implementation.
CHAR_WIDTHS = {
    " ": 3,
    "!": 5,
    '"': 6,
    "#": 7,
    "$": 7,
    "%": 11,
    "&": 10,
    "'": 3,
    "(": 5,
    ")": 5,
    "*": 7,
    "+": 7,
    ",": 4,
    "-": 5,
    ".": 4,
    "/": 6,
    "0": 7,
    "1": 7,
    "2": 7,
    "3": 7,
    "4": 7,
    "5": 7,
    "6": 7,
    "7": 7,
    "8": 7,
    "9": 7,
    ":": 4,
    ";": 4,
    "<": 7,
    "=": 7,
    ">": 7,
    "?": 7,
    "@": 13,
    "A": 9,
    "B": 8,
    "C": 8,
    "D": 9,
    "E": 7,
    "F": 7,
    "G": 9,
    "H": 9,
    "I": 4,
    "J": 5,
    "K": 8,
    "L": 6,
    "M": 12,
    "N": 10,
    "O": 10,
    "P": 8,
    "Q": 10,
    "R": 8,
    "S": 7,
    "T": 7,
    "U": 9,
    "V": 9,
    "W": 13,
    "X": 8,
    "Y": 7,
    "Z": 7,
    "[": 5,
    "\\": 6,
    "]": 5,
    "^": 7,
    "_": 7,
    "`": 4,
    "a": 7,
    "b": 8,
    "c": 6,
    "d": 8,
    "e": 8,
    "f": 5,
    "g": 7,
    "h": 8,
    "i": 4,
    "j": 4,
    "k": 7,
    "l": 4,
    "m": 12,
    "n": 8,
    "o": 8,
    "p": 8,
    "q": 8,
    "r": 5,
    "s": 6,
    "t": 5,
    "u": 8,
    "v": 7,
    "w": 11,
    "x": 7,
    "y": 7,
    "z": 6,
    "{": 5,
    "|": 7,
    "}": 5,
    "~": 7,
}

DEFAULT_COLUMN_WIDTH = 8.43
MAX_DIGIT_WIDTH = 7
CELL_PADDING = 5
CELL_BOUNDARY_PADDING = 7


def text_pixel_width(text: str) -> int:
    if "\n" in text:
        return max(text_pixel_width(segment) for segment in text.split("\n"))
    return sum(CHAR_WIDTHS.get(character, 8) for character in text)


def required_cell_width_pixels(display_text: str) -> int:
    if display_text == "":
        return 0
    return text_pixel_width(display_text) + CELL_BOUNDARY_PADDING


def column_width_to_pixels(width: float | int | None) -> int:
    numeric_width = DEFAULT_COLUMN_WIDTH if width is None else float(width)
    if numeric_width == 0:
        return 0
    if numeric_width < 1:
        return round(numeric_width * (MAX_DIGIT_WIDTH + CELL_PADDING))
    return round(numeric_width * MAX_DIGIT_WIDTH) + CELL_PADDING
