from pathlib import Path

import pytest

from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_csv_table import read_csv_table


def test_read_csv_table_reads_utf8_bom_and_records(tmp_path: Path) -> None:
    path = tmp_path / "source.csv"
    path.write_text("\ufeffReference,Amount\nINV-1,10\n", encoding="utf-8")

    result = read_csv_table(path)

    assert result == {
        "sheet": "",
        "headerRow": 1,
        "headers": ["reference", "amount"],
        "rows": [{"reference": "INV-1", "amount": "10"}],
        "rowCount": 1,
        "sourceRows": [2],
        "encoding": "utf-8-sig",
        "delimiter": ",",
        "truncated": False,
    }


def test_read_csv_table_skips_blank_rows_and_fills_trailing_cells(
    tmp_path: Path,
) -> None:
    path = tmp_path / "source.csv"
    path.write_text(
        "Reference,Amount,Note\nINV-1,10\n   , , \nINV-2,25,paid\n",
        encoding="utf-8",
    )

    result = read_csv_table(path)

    assert result["rows"] == [
        {"reference": "INV-1", "amount": "10", "note": ""},
        {"reference": "INV-2", "amount": "25", "note": "paid"},
    ]
    assert result["rowCount"] == 2


def test_read_csv_table_rejects_invalid_utf8(tmp_path: Path) -> None:
    path = tmp_path / "source.csv"
    path.write_bytes(b"Reference,Amount\nINV-1,\xff\n")

    with pytest.raises(SheetOpsError) as error:
        read_csv_table(path)

    assert error.value.error_code == "invalid_csv_encoding"
    assert error.value.details == {"path": str(path)}


def test_read_csv_table_rejects_empty_csv(tmp_path: Path) -> None:
    path = tmp_path / "source.csv"
    path.write_text("", encoding="utf-8")

    with pytest.raises(SheetOpsError) as error:
        read_csv_table(path)

    assert error.value.error_code == "empty_csv"
    assert error.value.details == {"path": str(path)}


def test_read_csv_table_uses_read_table_header_normalization(tmp_path: Path) -> None:
    path = tmp_path / "source.csv"
    path.write_text("Amount, amount \n10,20\n", encoding="utf-8")

    with pytest.raises(SheetOpsError) as error:
        read_csv_table(path)

    assert error.value.error_code == "invalid_headers"
    assert error.value.details == {"duplicates": ["amount"], "blankColumns": []}
