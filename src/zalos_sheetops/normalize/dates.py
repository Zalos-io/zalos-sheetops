from datetime import date, datetime

from zalos_sheetops.errors import SheetOpsError

_DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%b %d, %Y"]


def normalize_date(value: object) -> dict[str, object]:
    raw = "" if value is None else str(value).strip()
    if raw == "":
        return {"raw": raw, "normalized": ""}
    if isinstance(value, datetime):
        return {"raw": raw, "normalized": value.date().isoformat()}
    if isinstance(value, date):
        return {"raw": raw, "normalized": value.isoformat()}

    for date_format in _DATE_FORMATS:
        try:
            parsed = datetime.strptime(raw, date_format)
            return {"raw": raw, "normalized": parsed.date().isoformat()}
        except ValueError:
            continue
    raise SheetOpsError(
        "invalid_date", "Date value is not a supported date format", {"value": raw}
    )
