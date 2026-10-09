_ASCII_WHITESPACE = (" ", "\t", "\n", "\r", "\x0b", "\x0c")


def canonical_header(value: object) -> str:
    if value is None:
        return ""
    if type(value) is str:
        value = value.strip().lower()
        # " ".join(x.split()) only changes strings with inner whitespace, so
        # skip the split/join for the common single-token ASCII case. Non-ASCII
        # strings take the slow path because split() also breaks on unicode
        # whitespace such as non-breaking spaces.
        if value.isascii():
            for character in _ASCII_WHITESPACE:
                if character in value:
                    return " ".join(value.split())
            return value
        return " ".join(value.split())
    if type(value) is int or type(value) is float:
        # str() of a number never contains whitespace or uppercase letters.
        return str(value)
    return " ".join(str(value).strip().lower().split())
