from datetime import date, datetime

import pytest

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.normalize.dates import normalize_date


@pytest.mark.parametrize(
    ("raw", "normalized"),
    [
        ("2026-06-01", "2026-06-01"),
        ("06/01/2026", "2026-06-01"),
        ("6/1/26", "2026-06-01"),
        ("01-Jun-2026", "2026-06-01"),
        (date(2026, 6, 1), "2026-06-01"),
        (datetime(2026, 6, 1, 12, 30), "2026-06-01"),
        ("", ""),
        (None, ""),
    ],
)
def test_normalize_date_returns_iso_date(raw: object, normalized: str) -> None:
    assert normalize_date(raw) == {
        "raw": "" if raw is None else str(raw),
        "normalized": normalized,
    }


def test_normalize_date_rejects_unsupported_text() -> None:
    with pytest.raises(SheetOpsError) as error:
        normalize_date("not a date")

    assert error.value.error_code == "invalid_date"
