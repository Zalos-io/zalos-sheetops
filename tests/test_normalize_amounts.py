import pytest

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.normalize.amounts import normalize_amount


@pytest.mark.parametrize(
    ("raw", "normalized"),
    [
        ("$1,234.50", "1234.50"),
        ("USD 1,234.50", "1234.50"),
        ("1,234.50 USD", "1234.50"),
        ("(1,234.50)", "-1234.50"),
        ("-1,234.50", "-1234.50"),
        (1234.5, "1234.50"),
        ("", ""),
        (None, ""),
    ],
)
def test_normalize_amount_returns_decimal_string(raw: object, normalized: str) -> None:
    assert normalize_amount(raw) == {
        "raw": "" if raw is None else str(raw),
        "normalized": normalized,
    }


@pytest.mark.parametrize("raw", ["abc", "1.234,56", "$1,2,3", "1USD2", "$1$2"])
def test_normalize_amount_rejects_ambiguous_values(raw: str) -> None:
    with pytest.raises(SheetOpsError) as error:
        normalize_amount(raw)

    assert error.value.error_code == "invalid_amount"
