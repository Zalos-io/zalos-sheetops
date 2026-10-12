import json
from pathlib import Path

import pytest

from zalos_sheetops.cli import app
from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_csv_table import read_csv_table
from zalos_sheetops.testing import CliRunner


def test_csv_preview_preserves_ids_and_reports_source_rows(tmp_path: Path) -> None:
    path = tmp_path / "source.csv"
    path.write_text(
        "Reference\tAmount\tDate\n000123456789012345\t1,234.50\t2026-06-30\n"
        "\n0002\t25\t2026-07-01\n0003\t30\t2026-07-02\n",
        encoding="utf-16",
    )
    result = CliRunner().invoke(
        app,
        [
            "read-table",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    "encoding": "auto",
                    "delimiter": "\t",
                    "maxRows": 2,
                }
            ),
        ],
    )
    assert result.exit_code == 0, result.stdout
    data = json.loads(result.stdout)["data"]
    assert data["rows"] == [
        ["000123456789012345", "1,234.50", "2026-06-30"],
        ["0002", "25", "2026-07-01"],
    ]
    assert data["sourceRows"] == [2, 4]
    assert data["truncated"] is True
    assert data["encoding"] == "utf-16"


def test_csv_rejects_extra_fields_instead_of_discarding_them(tmp_path: Path) -> None:
    path = tmp_path / "source.csv"
    path.write_text("Reference,Amount\nINV-1,10,unexpected\n", encoding="utf-8")
    with pytest.raises(SheetOpsError) as caught:
        read_csv_table(path)
    assert caught.value.error_code == "invalid_csv_row"
    assert caught.value.details["row"] == 2


def test_reconcile_honors_csv_source_options(tmp_path: Path) -> None:
    left, right = tmp_path / "left.csv", tmp_path / "right.csv"
    left.write_text("Reference;Amount\n0001;12.50\n", encoding="utf-16")
    right.write_text("Reference,Amount\n0001,12.50\n", encoding="utf-8")
    payload = {
        "left": {"path": str(left), "encoding": "auto", "delimiter": ";"},
        "right": {"path": str(right)},
        "leftKeyColumns": ["reference"],
        "rightKeyColumns": ["reference"],
        "leftAmountColumn": "amount",
        "rightAmountColumn": "amount",
        "amountTolerance": "0.01",
    }
    result = CliRunner().invoke(app, ["reconcile", "--json", json.dumps(payload)])
    assert result.exit_code == 0, result.stdout
    assert json.loads(result.stdout)["data"]["summary"]["matched"] == 1


@pytest.mark.parametrize(
    "options", [{"encoding": 42}, {"delimiter": "||"}, {"encoding": "hex"}]
)
def test_csv_invalid_options_return_json(tmp_path: Path, options: dict) -> None:
    path = tmp_path / "source.csv"
    path.write_text("A,B\n1,2\n", encoding="utf-8")
    result = CliRunner().invoke(
        app,
        [
            "read-table",
            "--json",
            json.dumps(
                {
                    "path": str(path),
                    **options,
                }
            ),
        ],
    )
    assert result.exit_code == 1
    assert json.loads(result.stdout)["errorCode"] == "invalid_field"


def test_csv_rejects_null_bytes_from_wrong_decoding(tmp_path: Path) -> None:
    path = tmp_path / "source.csv"
    path.write_bytes("Reference,Amount\n001,10\n".encode("utf-16-le"))
    with pytest.raises(SheetOpsError) as caught:
        read_csv_table(path)
    assert caught.value.error_code == "invalid_csv_encoding"
