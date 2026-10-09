import json
from pathlib import Path

import pytest

from zalos_sheetops.benchmarks.public.agents import extract_usage, validate_resume


def test_usage_preserves_unknown_cost_status(tmp_path: Path) -> None:
    usage = tmp_path / "usage.json"
    usage.write_text(
        json.dumps(
            {
                "estimated_cost_usd": 0,
                "cost_status": "unknown",
                "cost_source": "none",
                "reasoning_tokens": 42,
                "total_including_auxiliary": {"total_tokens": 100},
                "secret": "must not be exported",
            }
        )
    )
    result = extract_usage(tmp_path, usage)
    assert result["cost_status"] == "unknown"
    assert result["reasoning_tokens"] == 42
    assert result["total_including_auxiliary"]["total_tokens"] == 100
    assert "secret" not in result


def test_resume_rejects_changed_model_but_allows_expanded_task_limit() -> None:
    old = {"model": "same", "taskIds": [1], "repeats": 1, "skillSha256": {"a": "x"}}
    validate_resume(old, {**old, "taskIds": [1, 2], "repeats": 3})
    with pytest.raises(ValueError, match="model"):
        validate_resume(old, {**old, "model": "different"})


def test_cleanup_selects_only_own_workspace_container(tmp_path: Path) -> None:
    from zalos_sheetops.benchmarks.public.agents import containers_for_workspace

    def container(identity: str, source: Path, label: str = "1") -> dict:
        return {
            "Id": identity,
            "Config": {"Labels": {"hermes-agent": label}},
            "Mounts": [{"Source": str(source), "Destination": "/workspace"}],
        }

    own = tmp_path / "own"
    assert containers_for_workspace(
        [
            container("own", own),
            container("other", tmp_path / "other"),
            container("unrelated", own, "0"),
        ],
        own,
    ) == ["own"]
