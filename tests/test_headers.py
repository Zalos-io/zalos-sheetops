from zalos_sheetops.io.headers import detect_header_row
from zalos_sheetops.normalize.text import canonical_header


def test_canonical_header_strips_lowercases_and_collapses_whitespace() -> None:
    assert canonical_header(" Charge   Amount USD ") == "charge amount usd"


def test_canonical_header_normalizes_non_string_values() -> None:
    assert canonical_header(None) == ""
    assert canonical_header(10) == "10"
    assert canonical_header(2.5) == "2.5"
    # bool must not take the int fast path: it lowercases like any other value
    assert canonical_header(True) == "true"
    assert canonical_header(False) == "false"


def test_canonical_header_collapses_unicode_whitespace() -> None:
    assert canonical_header("Charge\xa0Amount") == "charge amount"


def test_detect_header_row_skips_blank_title_and_merged_style_rows() -> None:
    rows = [
        [None, None, None],
        ["Merchant report", None, None],
        [" Date ", " Charge Amount USD ", "Reference"],
        ["2026-05-01", "12.34", "abc"],
    ]

    assert detect_header_row(rows, ["date", "charge amount usd"]) == 2


def test_detect_header_row_raises_when_expected_keywords_do_not_match_any_row() -> None:
    rows = [
        [None, None, None],
        ["Merchant report", None, None],
        ["Summary", "Totals", "Notes"],
    ]

    try:
        detect_header_row(rows, ["date", "charge amount usd"])
    except ValueError as exc:
        assert str(exc) == "Unable to detect header row"
    else:
        raise AssertionError(
            "detect_header_row() should fail closed when no keywords match"
        )


def test_detect_header_row_prefers_partial_keyword_match_over_wider_nonmatch() -> None:
    rows = [
        [None, None, None],
        [
            "Summary",
            "Totals",
            "Notes",
            "Extra",
            "More",
            "Fields",
            "Here",
            "Now",
            "Also",
            "Again",
            "Plus",
            "And",
            "Last",
        ],
        [" Date ", "Reference", None],
        ["2026-05-01", "abc", "xyz", "123"],
    ]

    assert detect_header_row(rows, ["date"]) == 2
