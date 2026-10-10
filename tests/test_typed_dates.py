from datetime import datetime, time, timedelta
from pathlib import Path

import pytest
import xlsxwriter

from zalos_sheetops.cli import app
from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_cells import read_cells
from zalos_sheetops.io.read_table import read_table
from zalos_sheetops.testing import CliRunner


@pytest.mark.parametrize("date_1904", [False, True])
def test_typed_xml_dates_match_calendar_values_and_leave_raw_serials_available(
    tmp_path: Path,
    date_1904: bool,
) -> None:
    path = tmp_path / "dates.xlsx"
    with xlsxwriter.Workbook(path, {"date_1904": date_1904}) as book:
        sheet = book.add_worksheet("Dates")
        sheet.write_row(0, 0, ["Date", "Time", "Duration", "Amount", "Formula date"])
        date_format = book.add_format({"num_format": "yyyy-mm-dd hh:mm:ss.000"})
        sheet.write_datetime(1, 0, datetime(2026, 6, 15, 12, 30), date_format)
        sheet.write_datetime(
            1, 1, time(6, 30), book.add_format({"num_format": "hh:mm"})
        )
        sheet.write_datetime(
            1, 2, timedelta(hours=36), book.add_format({"num_format": "[h]:mm:ss"})
        )
        sheet.write_number(1, 3, 123.45, book.add_format({"num_format": '0.00 "days"'}))
        sheet.write_formula(
            1,
            4,
            "=A2",
            date_format,
            44726.52083333333 if date_1904 else 46188.52083333333,
        )
    typed = read_cells(
        path,
        "Dates",
        "A2:E2",
        ["values", "cachedValues", "formulas"],
        engine="xml",
        value_mode="typed",
    )
    cells = {cell["address"]: cell for cell in typed["cells"]}
    assert cells["A2"]["value"] == "2026-06-15T12:30:00"
    assert cells["B2"]["value"] == "06:30:00"
    assert cells["C2"]["value"] == "1 day, 12:00:00"
    assert cells["D2"]["value"] == 123.45
    assert cells["E2"]["cachedValue"] == "2026-06-15T12:30:00"
    assert cells["E2"]["formula"] == "=A2"
    raw = read_cells(path, "Dates", "A2:A2", ["values"], engine="xml")
    assert raw["cells"][0]["value"] == pytest.approx(
        44726.52083333333 if date_1904 else 46188.52083333333
    )
    table = read_table(path, "Dates", "row", 1, value_mode="typed")
    assert table["rows"] == [
        [
            "2026-06-15 12:30:00",
            "06:30:00",
            "1 day, 12:00:00",
            "123.45",
            "2026-06-15 12:30:00",
        ]
    ]


@pytest.mark.parametrize("command", [["cells", "read"], ["read-table"]])
def test_read_commands_reject_unknown_value_mode(
    command: list[str], tmp_path: Path
) -> None:
    import json

    request = {
        "path": str(tmp_path / "unused.xlsx"),
        "sheet": "Sheet",
        "range": "A1",
        "valueMode": "guess",
    }
    result = CliRunner().invoke(
        app, [*command, "--json", json.dumps(request), "--raw-output"]
    )
    assert result.exit_code == 1
    assert json.loads(result.output)["errorCode"] == "invalid_field"
    assert json.loads(result.output)["details"]["field"] == "valueMode"


def test_typed_mode_rejects_unknown_mode_before_opening_workbook(
    tmp_path: Path,
) -> None:
    with pytest.raises(SheetOpsError) as error:
        read_cells(
            tmp_path / "unused.xlsx", "Sheet", "A1", engine="xml", value_mode="guess"
        )
    assert error.value.error_code == "invalid_field"
