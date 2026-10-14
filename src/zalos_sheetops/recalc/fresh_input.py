"""Remove direct formula result nodes without rewriting worksheet XML."""

from xml.parsers import expat

from zalos_sheetops.io.xlsx_xml import MAIN_NS


def _opening_end(xml: bytes, index: int, encoding: str) -> int:
    quote = 0
    width = 2 if encoding.startswith("utf-16") else 1
    for position in range(index, len(xml), width):
        value = ord(xml[position : position + width].decode(encoding))
        if quote:
            if value == quote:
                quote = 0
        elif value in {34, 39}:
            quote = value
        elif value == 62:
            return position + width
    raise ValueError("Unterminated XML opening tag")


def without_formula_caches(xml: bytes) -> bytes:
    encoding = (
        "utf-16-le"
        if xml.startswith((b"\xff\xfe", b"<\x00"))
        else "utf-16-be"
        if xml.startswith((b"\xfe\xff", b"\x00<"))
        else "latin1"
    )
    closing = "/>".encode(encoding)
    parser = expat.ParserCreate(namespace_separator="}")
    tags = {name: MAIN_NS + "}" + name for name in ("c", "f", "v", "is")}
    depth, cell_depth, cache_start, cache_stop = 0, 0, 0, 0
    has_formula = False
    caches, removals = [], []

    def start(name: str, attributes: dict[str, str]) -> None:
        nonlocal depth, cell_depth, has_formula, caches, cache_start, cache_stop
        depth += 1
        if name == tags["c"]:
            cell_depth, has_formula, caches = depth, False, []
        elif cell_depth and depth == cell_depth + 1:
            if name == tags["f"]:
                has_formula = True
            elif name in {tags["v"], tags["is"]}:
                cache_start = parser.CurrentByteIndex
                opening_end = _opening_end(xml, cache_start, encoding)
                cache_stop = (
                    opening_end
                    if xml[opening_end - len(closing) : opening_end] == closing
                    else 0
                )

    def end(name: str) -> None:
        nonlocal depth, cell_depth
        if cell_depth and depth == cell_depth + 1 and name in {tags["v"], tags["is"]}:
            index = parser.CurrentByteIndex
            stop = cache_stop or _opening_end(xml, index, encoding)
            caches.append((cache_start, stop))
        elif cell_depth == depth and name == tags["c"]:
            if has_formula:
                removals.extend(caches)
            cell_depth = 0
        depth -= 1

    parser.StartElementHandler, parser.EndElementHandler = start, end
    parser.Parse(xml, True)
    pieces, cursor = [], 0
    for start_index, stop in removals:
        pieces.append(xml[cursor:start_index])
        cursor = stop
    pieces.append(xml[cursor:])
    return b"".join(pieces)
