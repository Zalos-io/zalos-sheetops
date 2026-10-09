from zalos_sheetops.display.widths import (
    column_width_to_pixels,
    required_cell_width_pixels,
    text_pixel_width,
)


def test_text_width_uses_xlsxwriter_character_widths() -> None:
    assert text_pixel_width("Hello") == 33
    assert text_pixel_width("12345") == 35
    assert text_pixel_width("界") == 8


def test_required_cell_width_adds_excel_padding() -> None:
    assert required_cell_width_pixels("") == 0
    assert required_cell_width_pixels("Hello") == 40
    assert required_cell_width_pixels("12/31/2026") == 75


def test_column_width_to_pixels_matches_excel_character_width_conversion() -> None:
    assert column_width_to_pixels(None) == 64
    assert column_width_to_pixels(0) == 0
    assert column_width_to_pixels(0.5) == 6
    assert column_width_to_pixels(8.43) == 64
    assert column_width_to_pixels(10) == 75
