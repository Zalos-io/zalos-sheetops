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
