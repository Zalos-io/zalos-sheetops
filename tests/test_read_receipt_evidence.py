import inspect
import json
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904

from zalos_sheetops.cli import app
from zalos_sheetops.io.read_cells import read_cells
from zalos_sheetops.io.read_csv_table import read_csv_table
from zalos_sheetops.io.read_table import read_table
from zalos_sheetops.reconcile.exact import reconcile_exact
from zalos_sheetops.testing import CliRunner


@pytest.mark.parametrize("reader", [read_table, read_csv_table, reconcile_exact])
def test_public_api_doc_lookup_is_usable_without_reading_implementation(
    reader: object,
) -> None:
    # This is the standard lookup that failed in the captured agent commands.
    excerpt = inspect.getdoc(reader)[:2000]
    assert isinstance(excerpt, str) and excerpt.strip()


@pytest.mark.parametrize("date_1904", [False, True])
def test_typed_receipts_expose_epoch_and_keep_raw_contract(
    tmp_path: Path, date_1904: bool
) -> None:
    path = tmp_path / "source.xlsx"
    book = Workbook()
    if date_1904:
        book.epoch = CALENDAR_MAC_1904
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["Reference", "Amount", "Date"])
    sheet.append(
        ["00001", "0.12345678901234567890123456789012345", datetime(2026, 6, 30)]
    )
    book.save(path)
    expected = "1904" if date_1904 else "1900"
    typed_results = [
        read_table(path, "Data", "row", 1, value_mode="typed"),
        read_cells(
            path,
            "Data",
            "A1:C2",
            ["values", "dataTypes", "numberFormats"],
            engine="xml",
            value_mode="typed",
        ),
    ]
    for typed in typed_results:
        assert typed["dateSystem"] == expected and typed["valueMode"] == "typed"
    raw_results = [
        read_table(path, "Data", "row", 1),
        read_cells(path, "Data", "A1:C2", engine="xml"),
    ]
    for raw in raw_results:
        assert "dateSystem" not in raw and "valueMode" not in raw
    table = read_table(path, "Data", "row", 1, value_mode="typed")
    assert table["rows"][0][:2] == ["00001", "0.12345678901234567890123456789012345"]
    assert (
        datetime.fromisoformat(table["rows"][0][2]).date().isoformat() == "2026-06-30"
    )
    cells = read_cells(
        path,
        "Data",
        "C2",
        ["values", "numberFormats"],
        engine="xml",
        value_mode="typed",
    )
    assert cells["cells"][0]["value"] == "2026-06-30T00:00:00"
    assert cells["cells"][0]["numberFormat"] == "yyyy-mm-dd h:mm:ss"


def test_profile_cli_keeps_typed_date_evidence(tmp_path: Path) -> None:
    path = tmp_path / "source.xlsx"
    book = Workbook()
    book.active.append(["date"])
    book.active.append([datetime(2026, 6, 30)])
    book.save(path)
    result = CliRunner().invoke(
        app,
        [
            "profile-table",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "headerStrategy": "row",
                    "headerRow": 1,
                    "valueMode": "typed",
                }
            ),
        ],
    )
    assert result.exit_code == 0
    data = json.loads(result.output)["data"]
    assert data["source"]["dateSystem"] == "1900"
    assert data["source"]["valueMode"] == "typed"
    assert data["profile"]["validated"] is True
