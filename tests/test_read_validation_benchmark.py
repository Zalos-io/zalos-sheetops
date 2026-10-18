from pathlib import Path

from zalos_sheetops.benchmarks.read_validation import grade


def test_benchmark_rejects_silent_success_on_invalid_input() -> None:
    assert grade({"errorCode": "invalid_key_value"}, {"data": {"matched": 1}}) is False
    assert (
        grade({"errorCode": "invalid_key_value"}, {"errorCode": "invalid_key_value"})
        is True
    )


def test_benchmark_requires_exact_expected_values() -> None:
    assert grade({"data": {"id": "0001"}}, {"data": {"id": "1"}}) is False
    assert grade({"data": {"id": "0001"}}, {"data": {"id": "0001"}}) is True


def test_prepared_typed_case_uses_documented_string_table_values(
    tmp_path: Path,
) -> None:
    import json

    from zalos_sheetops.benchmarks.read_validation import prepare, worker

    root = tmp_path / "fixtures"
    prepare(root)
    manifest = json.loads((root / "manifest.json").read_text())
    case = next(row for row in manifest["cases"] if row["id"] == "development-typed")
    result = worker(case, root)
    assert result["passed"], result["observed"]
    assert Path(result["packageSource"]).name == "__init__.py"


def test_run_rejects_missing_source_instead_of_using_installed_package(
    tmp_path: Path,
) -> None:
    import json

    import pytest

    from zalos_sheetops.benchmarks.read_validation import run

    root = tmp_path / "fixtures"
    root.mkdir()
    case = {
        "id": "totals",
        "split": "development",
        "operation": "totals",
        "left": ["1"],
        "right": ["1"],
        "expected": {"data": {"validated": True}},
    }
    cases = [
        {**case, "id": split, "split": split}
        for split in ["development", "reserved", "performance"]
    ]
    (root / "manifest.json").write_text(json.dumps({"files": {}, "cases": cases}))
    with pytest.raises(ValueError, match="Source"):
        run(root, tmp_path / "missing", tmp_path / "result.json", 1)
