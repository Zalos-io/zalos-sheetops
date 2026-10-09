import re
from decimal import Decimal, InvalidOperation

from zalos_sheetops.errors import SheetOpsError

_AMOUNT_RE = re.compile(r"^-?\d+(\.\d+)?$")
_COMMA_AMOUNT_RE = re.compile(r"^-?\d{1,3}(,\d{3})+(\.\d+)?$")


def normalize_amount(value: object) -> dict[str, object]:
    raw = "" if value is None else str(value).strip()
    if raw == "":
        return {"raw": raw, "normalized": ""}

    negative = raw.startswith("(") and raw.endswith(")")
    cleaned = raw.strip("()").strip()
    cleaned = _strip_currency_token(cleaned, "$")
    cleaned = _strip_currency_token(cleaned, "USD")
    if cleaned.startswith("+"):
        cleaned = cleaned[1:]
    if "," in cleaned and not _COMMA_AMOUNT_RE.match(cleaned):
        raise SheetOpsError(
            "invalid_amount", "Amount value is not a supported decimal", {"value": raw}
        )
    cleaned = cleaned.replace(",", "")

    if not _AMOUNT_RE.match(cleaned):
        raise SheetOpsError(
            "invalid_amount", "Amount value is not a supported decimal", {"value": raw}
        )

    try:
        amount = Decimal(cleaned)
    except InvalidOperation as error:
        raise SheetOpsError(
            "invalid_amount", "Amount value is not a supported decimal", {"value": raw}
        ) from error

    if negative:
        amount = -amount
    normalized = f"{amount.quantize(Decimal('0.01'))}"
    return {"raw": raw, "normalized": normalized}


def _strip_currency_token(value: str, token: str) -> str:
    cleaned = value.strip()
    if cleaned.startswith(token):
        return cleaned[len(token) :].strip()
    if cleaned.endswith(token):
        return cleaned[: -len(token)].strip()
    return cleaned
