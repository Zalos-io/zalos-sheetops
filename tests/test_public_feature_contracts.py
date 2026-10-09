import json
from pathlib import Path

from openpyxl import Workbook

from zalos_sheetops.benchmarks.public.contracts import run_contracts


def test_feature_suite_checks_metadata_and_formats(tmp_path: Path) -> None:
    root = tmp_path / "data" / "excelbench"
    root.mkdir(parents=True)
    book = Workbook()
    book.active.title = "formats"
    book.active["B2"] = 12.3
    book.active["B2"].number_format = "$0.00"
    book.active.merge_cells("B3:D3")
    book.save(root / "formats.xlsx")
    (root / "manifest.json").write_text(
        json.dumps(
            {
                "files": [
                    {
                        "path": "formats.xlsx",
                        "feature": "number_formats",
                        "test_cases": [
                            {"row": 2, "expected": {"number_format": '"$"0.00'}}
                        ],
                    }
                ]
            }
        )
    )
    records = run_contracts(tmp_path / "data", tmp_path / "results")
    assert {r["case"] for r in records} >= {"metadata-number_formats", "number_formats"}
    assert all(r["status"] == "pass" for r in records)
    assert all(r["seconds"] > 0 for r in records)
    style = next(r for r in records if r["case"] == "number_formats")
    assert style["fixtureManifestDifference"] == {
        "manifest": '"$"0.00',
        "stored": "$0.00",
    }
