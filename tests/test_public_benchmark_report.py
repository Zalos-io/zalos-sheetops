from zalos_sheetops.benchmarks.public.report import summarize_agents


def test_failed_pairs_cannot_enter_speed_comparison() -> None:
    records = [
        {
            "task": 1,
            "arm": "baseline",
            "repeat": 0,
            "seconds": 100,
            "grading": {"status": "incorrect"},
            "usage": {"cost_status": "unknown"},
        },
        {
            "task": 1,
            "arm": "sheetops",
            "repeat": 0,
            "seconds": 1,
            "grading": {"status": "pass"},
            "usage": {"cost_status": "unknown"},
        },
    ]
    result = summarize_agents(records)
    assert result["correctPairCount"] == 0
    assert result["medianPairedSpeedup"] is None
    assert result["arms"]["baseline"]["statusCounts"] == {"incorrect": 1}
    assert result["arms"]["baseline"]["costAvailable"] is False


def test_pairing_matches_task_and_repeat_and_keeps_cost_unavailable() -> None:
    records = [
        {
            "task": 1,
            "arm": "baseline",
            "repeat": 0,
            "seconds": 10,
            "grading": {"status": "pass"},
            "usage": {"cost_status": "unknown"},
        },
        {
            "task": 1,
            "arm": "sheetops",
            "repeat": 0,
            "seconds": 5,
            "grading": {"status": "pass"},
            "usage": {"cost_status": "unknown"},
        },
        {
            "task": 2,
            "arm": "baseline",
            "repeat": 0,
            "seconds": 1,
            "grading": {"status": "pass"},
            "usage": {},
        },
    ]
    result = summarize_agents(records)
    assert result["correctPairCount"] == 1
    assert result["medianPairedSpeedup"] == 2
    assert result["completePairCount"] == 1
    assert result["attemptCount"] == 3


def test_unknown_reader_failures_are_not_mislabeled_as_date_gaps() -> None:
    from zalos_sheetops.benchmarks.public.report import failure_catalog

    records = [{"adapter": "sheetops-cli", "case": "example", "status": "error"}]
    failures = failure_catalog(records, [])
    assert failures[0]["id"] == "sheetops-operation-contract"


def test_report_uses_explicit_regraded_records_and_requires_operations(
    tmp_path,
) -> None:
    import json

    from zalos_sheetops.benchmarks.public.report import build_report

    agent_root = tmp_path / "agents"
    agent_root.mkdir()
    (agent_root / "agent-metadata.json").write_text(
        json.dumps({"taskIds": [1], "repeats": 1})
    )
    records = [
        {
            "task": 1,
            "arm": arm,
            "repeat": 0,
            "seconds": 1,
            "grading": {"status": "pass"},
            "usage": {},
        }
        for arm in ["baseline", "sheetops"]
    ]
    override = tmp_path / "graded.jsonl"
    override.write_text("".join(json.dumps(r) + "\n" for r in records))
    result = build_report(tmp_path, tmp_path / "report", agent_records=override)
    assert result["agents"]["correctPairCount"] == 1
    assert result["status"] == "in_progress"


def test_duplicate_attempts_cannot_satisfy_complete_coverage() -> None:
    from zalos_sheetops.benchmarks.public.report import coverage

    records = [{"case": "a", "adapter": "openpyxl", "repeat": 0}] * 2
    expected = {("a", "openpyxl", 0), ("b", "openpyxl", 0)}
    result = coverage(records, expected, ("case", "adapter", "repeat"))
    assert result["complete"] is False
    assert result["duplicates"] == 1
    assert result["missing"] == 1


def test_all_failure_components_survive_diagnostic_cap_and_report_adoption() -> None:
    from zalos_sheetops.benchmarks.public.report import failure_catalog

    records = [
        {
            "task": "123",
            "arm": "sheetops",
            "repeat": 0,
            "seconds": 8,
            "grading": {
                "status": "incorrect",
                "failureCounts": {"values": 1, "styles": 50, "preservation": 2},
            },
            "usage": {
                "sheetopsSuccessfulResponses": 4,
                "sheetopsFailedResponses": 1,
                "sheetopsCommandCounts": {"inspect": 3, "cells": 2},
                "sheetopsErrorCounts": {"sheet_not_found": 1},
            },
        }
    ]
    failures = failure_catalog([], records)
    assert {r["id"] for r in failures} == {
        "answer-values",
        "explicit-styles",
        "workbook-preservation",
        "tool-contract-misuse",
    }
    assert all(r["examples"][0]["task"] == "123" for r in failures)
    summary = summarize_agents(records)["arms"]["sheetops"]
    assert summary["sheetopsAdoption"]["tasksWithResponses"] == 1
    assert summary["sheetopsAdoption"]["successfulResponses"] == 4
    assert summary["sheetopsAdoption"]["errorCounts"] == {"sheet_not_found": 1}


def test_operation_table_never_displays_incorrect_output_timing() -> None:
    from zalos_sheetops.benchmarks.public.report import operation_table

    lines = operation_table(
        [
            {
                "case": "example",
                "adapter": "openpyxl",
                "status": "pass",
                "seconds": 2,
                "peakProcessRssBytes": 1048576,
            },
            {
                "case": "example",
                "adapter": "sheetops-cli",
                "status": "incorrect",
                "seconds": 0.001,
            },
        ]
    )
    assert "2.000" in "\n".join(lines)
    assert "incorrect" in "\n".join(lines)
    assert "0.001" not in "\n".join(lines)


def test_recovered_tool_errors_are_separate_from_task_failures() -> None:
    from zalos_sheetops.benchmarks.public.report import failure_catalog

    record = {
        "task": 1,
        "arm": "sheetops",
        "repeat": 0,
        "grading": {"status": "pass"},
        "usage": {
            "sheetopsFailedResponses": 1,
            "sheetopsErrorCounts": {"invalid_field": 1},
        },
    }
    failures = failure_catalog([], [record])
    assert len(failures) == 1
    assert failures[0]["id"] == "tool-contract-misuse"
    assert failures[0]["errorCounts"] == {"invalid_field": 1}
    assert failures[0]["kind"] == "observed tool errors; task may still pass"


def test_oracle_coverage_is_separate_from_scored_attempts_and_formula_errors() -> None:
    from zalos_sheetops.benchmarks.public.report import failure_catalog

    records = [
        {
            "task": 1,
            "arm": "baseline",
            "repeat": 0,
            "seconds": 1,
            "usage": {},
            "grading": {"status": "oracle_preservation_conflict"},
        },
        {
            "task": 2,
            "arm": "baseline",
            "repeat": 0,
            "seconds": 1,
            "usage": {},
            "grading": {
                "status": "incorrect",
                "failureCounts": {"values": 1},
                "valueDifferences": [
                    {"cell": "A1", "expected": 2, "actual": "#VALUE!"}
                ],
            },
        },
    ]
    summary = summarize_agents(records)
    assert summary["arms"]["baseline"]["scorableAttempts"] == 1
    assert summary["arms"]["baseline"]["unscorableAttempts"] == 1
    failures = failure_catalog([], records)
    assert {r["id"] for r in failures} == {"oracle-coverage", "formula-evaluation"}
