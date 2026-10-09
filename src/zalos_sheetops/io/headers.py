from zalos_sheetops.normalize.text import canonical_header


def detect_header_row(
    rows: list[list[object]],
    expected_keywords: list[str] | None = None,
) -> int:
    expected = {canonical_header(keyword) for keyword in expected_keywords or []}
    best_index = -1
    best_score = -1

    for index, row in enumerate(rows):
        headers = [canonical_header(value) for value in row]
        non_blank = {header for header in headers if header}
        if not non_blank:
            continue
        overlap = len(non_blank & expected)
        if expected and overlap == 0:
            continue
        score = overlap * 100 + len(non_blank)
        if score > best_score:
            best_score = score
            best_index = index

    if best_index == -1:
        raise ValueError("Unable to detect header row")
    return best_index
