from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.normalize.preview import preview_normalization


def test_preview_normalization_returns_row_level_values_and_issues(
    tmp_path: Path,
) -> None:
    path = tmp_path / "source.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["Amount", "Date"])
    sheet.append(["$1,234.50", "06/01/2026"])
    sheet.append(["abc", "not a date"])
    workbook.save(path)

    result = preview_normalization(
        path,
        "Data",
        "detect",
        None,
        [
            {"source": "Amount", "kind": "amount", "name": "amount"},
            {"source": "Date", "kind": "date", "name": "date"},
        ],
        None,
    )

    assert result["rowCount"] == 2
    assert result["issueCount"] == 2
    assert result["previewRows"][0]["values"]["amount"]["normalized"] == "1234.50"
    assert result["previewRows"][0]["values"]["date"]["normalized"] == "2026-06-01"
    assert result["previewRows"][1]["issues"] == [
        {"column": "Amount", "kind": "amount", "errorCode": "invalid_amount"},
        {"column": "Date", "kind": "date", "errorCode": "invalid_date"},
    ]


def test_preview_normalization_pushes_max_rows_into_table_read(
    tmp_path: Path,
    monkeypatch: object,
) -> None:
    path = tmp_path / "source.xlsx"
    seen: dict[str, object] = {}

    def fake_read_table(
        table_path: Path,
        sheet_name: str | None,
        header_strategy: str,
        header_row: int | None,
        max_rows: int | None = None,
    ) -> dict[str, object]:
        seen["maxRows"] = max_rows
        return {
            "sheet": sheet_name,
            "headerRow": 1,
            "headers": ["amount"],
            "rows": [["$1.00"]],
            "rowCount": 1000,
            "returnedRowCount": 1,
            "truncated": True,
        }

    monkeypatch.setattr(
        "zalos_sheetops.normalize.preview.read_table",
        fake_read_table,
    )

    result = preview_normalization(
        path,
        "Data",
        "row",
        1,
        [{"source": "Amount", "kind": "amount", "name": "amount"}],
        max_rows=1,
    )

    assert seen["maxRows"] == 1
    assert result["rowCount"] == 1000
    assert len(result["previewRows"]) == 1
