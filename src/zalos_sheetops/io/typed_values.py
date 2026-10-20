"""Opt-in Excel date conversion without changing the raw XML read contract."""

from datetime import date, datetime, time, timedelta

from zalos_sheetops.errors import SheetOpsError


def validate_value_mode(value_mode: object) -> str:
    if value_mode not in ("raw", "typed"):
        raise SheetOpsError(
            "invalid_field",
            "Field has invalid type",
            {
                "field": "valueMode",
                "expected": "raw | typed",
            },
        )
    return value_mode


class TypedValues:
    def __init__(self, number_formats: dict[int, str], date_1904: bool) -> None:
        from openpyxl.styles.numbers import is_date_format, is_timedelta_format
        from openpyxl.utils.datetime import (
            CALENDAR_MAC_1904,
            CALENDAR_WINDOWS_1900,
            from_excel,
            from_ISO8601,
        )

        self.epoch = CALENDAR_MAC_1904 if date_1904 else CALENDAR_WINDOWS_1900
        self.date_system = "1904" if date_1904 else "1900"
        self._from_excel = from_excel
        self._from_iso = from_ISO8601
        self.date_styles = {
            style: is_timedelta_format(fmt)
            for style, fmt in number_formats.items()
            if is_date_format(fmt)
        }

    def convert(
        self, value: object, style: str | None, data_type: str | None
    ) -> object:
        if data_type == "d" and isinstance(value, str):
            return self._from_iso(value)
        if data_type not in (None, "n") or type(value) not in (int, float):
            return value
        style_id = int(style or 0)
        if style_id not in self.date_styles:
            return value
        try:
            return self._from_excel(
                value, self.epoch, timedelta=self.date_styles[style_id]
            )
        except (OverflowError, ValueError) as error:
            raise SheetOpsError(
                "invalid_date",
                "Excel date is outside the supported range",
                {"value": value, "style": style_id},
            ) from error


def json_typed_value(value: object) -> object:
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return str(value)
    return value
