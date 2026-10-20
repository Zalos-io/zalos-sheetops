"""Rejected records provide bounded source evidence without another file read."""

import json
from pathlib import Path

import pytest

from zalos_sheetops.cli import app
from zalos_sheetops.errors import SheetOpsError
from zalos_sheetops.io.read_csv_table import read_csv_table
from zalos_sheetops.testing import CliRunner
from zalos_sheetops.validation.table_profile import profile_table


def complete_table(
    rows: list[dict[str, object]], source_rows: list[int] | None = None
) -> dict[str, object]:
    return {
        "headers": ["document", "value"],
        "rows": rows,
        "rowCount": len(rows),
        "returnedRowCount": len(rows),
        "truncated": False,
        "sourceRows": source_rows or list(range(2, len(rows) + 2)),
    }


def rejected(table: dict[str, object]) -> SheetOpsError:
    with pytest.raises(SheetOpsError) as caught:
        profile_table(table, key_columns=["document"], amount_columns=["value"])
    assert caught.value.error_code == "invalid_table_values"
    return caught.value


def test_rejection_shows_exact_small_source_values_and_keeps_record_dedup() -> None:
    table = complete_table([{"document": " \t", "value": "NaN"}], [9])
    before = json.dumps(table)
    details = rejected(table).details
    assert details["invalidRecordCount"] == 1
    assert details["issueCount"] == 2
    assert details["issues"] == [
        {
            "record": 1,
            "sourceRow": 9,
            "column": "document",
            "errorCode": "invalid_key_value",
            "valueType": "str",
            "valuePreview": " \t",
            "valueTruncated": False,
        },
        {
            "record": 1,
            "sourceRow": 9,
            "column": "value",
            "errorCode": "invalid_amount_value",
            "valueType": "str",
            "valuePreview": "NaN",
            "valueTruncated": False,
        },
    ]
    assert json.dumps(table) == before


def test_large_invalid_value_has_explicit_bounded_preview() -> None:
    value = "bad amount 東京\n" * 100
    details = rejected(complete_table([{"document": "0001", "value": value}])).details
    issue = details["issues"][0]
    assert issue["valuePreview"] == value[:80]
    assert issue["valueTruncated"] is True
    assert issue["valueType"] == "str"
    assert details["invalidRecordCount"] == 1


def test_preview_limit_never_limits_complete_invalid_population_count() -> None:
    rows = [{"document": f"000{i}", "value": "-Infinity"} for i in range(1001)]
    details = rejected(complete_table(rows)).details
    assert details["invalidRecordCount"] == details["issueCount"] == 1001
    assert len(details["issues"]) == 20
    assert details["issuesTruncated"] is True
    assert all(issue["valuePreview"] == "-Infinity" for issue in details["issues"])


@pytest.mark.parametrize("value", [float("nan"), float("inf"), None, False])
def test_non_string_rejection_evidence_is_json_safe(value: object) -> None:
    details = rejected(complete_table([{"document": "0001", "value": value}])).details
    assert details["invalidRecordCount"] == 1
    assert all(
        issue["valueType"] == type(value).__name__ for issue in details["issues"]
    )
    assert all("valuePreview" not in issue for issue in details["issues"])
    json.dumps(details, allow_nan=False)


def test_unsupported_objects_are_not_stringified_for_evidence() -> None:
    class NoDisplay:
        def __str__(self) -> str:
            raise AssertionError("Unsupported values must not be coerced")

        def __repr__(self) -> str:
            raise AssertionError("Unsupported values must not be displayed")

    details = rejected(
        complete_table([{"document": "0001", "value": NoDisplay()}])
    ).details
    assert details["invalidRecordCount"] == 1
    assert all(issue["valueType"] == "NoDisplay" for issue in details["issues"])
    json.dumps(details, allow_nan=False)


def test_multiline_rejections_keep_physical_source_rows_and_decoded_values(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.csv"
    source.write_text(
        'document,value\n0001,"not\na decimal"\n0002,NaN\n', encoding="utf-16"
    )
    before = source.read_bytes()
    table = read_csv_table(source, encoding="auto", delimiter=",", max_rows=None)
    details = rejected(table).details
    assert details["invalidRecordCount"] == 2
    assert [issue["sourceRow"] for issue in details["issues"]] == [2, 4]
    assert [issue["valuePreview"] for issue in details["issues"]] == [
        "not\na decimal",
        "NaN",
    ]
    assert source.read_bytes() == before


def test_profile_cli_returns_rejection_evidence_with_nonzero_exit(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.csv"
    source.write_text("document,value\n,NaN\n", encoding="utf-8")
    result = CliRunner().invoke(
        app,
        [
            "profile-table",
            "--json",
            json.dumps(
                {
                    "path": str(source),
                    "keyColumns": ["document"],
                    "amountColumns": ["value"],
                }
            ),
        ],
    )
    assert result.exit_code == 1
    payload = json.loads(result.output)
    assert payload["ok"] is False
    assert payload["errorCode"] == "invalid_table_values"
    assert payload["details"]["invalidRecordCount"] == 1
    assert [issue["valuePreview"] for issue in payload["details"]["issues"]] == [
        "",
        "NaN",
    ]
