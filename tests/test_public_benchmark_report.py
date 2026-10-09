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
