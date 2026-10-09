import json
from pathlib import Path

from zalos_sheetops.benchmarks.public.scoring import score_outputs, summarize
from zalos_sheetops.benchmarks.public.tasks import select_tasks


def test_wrong_cell_never_receives_speed_ranking() -> None:
    expected = [[True, 1, "001", None]]
    result = score_outputs(expected, [[1, True, 1, ""]])
    assert result["correct"] is False
    assert result["mismatchCount"] == 4
    summary = summarize(
        [
            {"case": "a", "adapter": "bad", "status": "incorrect", "seconds": 0.01},
            {"case": "a", "adapter": "good", "status": "pass", "seconds": 1.0},
            {"case": "a", "adapter": "good", "status": "timeout", "seconds": 2.0},
        ]
    )
    assert summary[0]["medianSeconds"] is None
    assert summary[1]["passRate"] == 0.5
    assert summary[1]["medianSeconds"] is None


def test_score_checks_missing_rows_and_records_first_difference() -> None:
    result = score_outputs([[1], [2]], [[1]])
    assert result["correct"] is False
    assert result["firstDifferences"][0]["row"] == 2
    assert score_outputs([[1.25]], [[1.25]])["correct"] is True


def test_task_selection_is_seeded_and_does_not_mutate_source(tmp_path: Path) -> None:
    tasks = [
        {"id": str(n), "instruction_type": "Cell-Level Manipulation"} for n in range(40)
    ]
    source = tmp_path / "dataset.json"
    source.write_text(json.dumps(tasks), encoding="utf-8")
    first = select_tasks(tasks, count=30, seed=20261009)
    assert first == select_tasks(tasks, count=30, seed=20261009)
    assert len({task["id"] for task in first}) == 30
    assert json.loads(source.read_text(encoding="utf-8")) == tasks
