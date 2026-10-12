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


def test_usage_counts_sheetops_responses_without_exporting_payloads(
    tmp_path: Path,
) -> None:
    import sqlite3

    responses = [
        {"ok": True, "command": "inspect", "data": {"private_payload": "exclude"}},
        {
            "ok": False,
            "command": "cells read",
            "errorCode": "invalid_field",
            "message": "exclude",
        },
    ]
    with sqlite3.connect(tmp_path / "state.db") as db:
        db.execute("CREATE TABLE messages(role TEXT, content TEXT)")
        output = "\n".join(json.dumps(r) for r in responses)
        db.execute(
            "INSERT INTO messages VALUES('tool', ?)", (json.dumps({"output": output}),)
        )
        db.execute(
            "INSERT INTO messages VALUES('tool', ?)",
            (json.dumps({"output": "ordinary output"}),),
        )
    result = extract_usage(tmp_path, tmp_path / "absent.json")
    assert result["observedToolCalls"] == 2
    assert result["sheetopsSuccessfulResponses"] == 1
    assert result["sheetopsFailedResponses"] == 1
    assert result["sheetopsErrorCounts"] == {"invalid_field": 1}
    assert result["sheetopsCommandCounts"] == {"inspect": 1, "cells read": 1}
    assert "exclude" not in json.dumps(result)


def test_benchmark_counts_only_the_approved_delivered_bytes(tmp_path: Path) -> None:
    import hashlib

    from zalos_sheetops.benchmarks.public.agents import delivery_status

    output = tmp_path / "output.xlsx"
    output.write_bytes(b"candidate")
    assert delivery_status(tmp_path, output)["status"] == "blocked"
    receipt = {str(output.resolve()): hashlib.sha256(output.read_bytes()).hexdigest()}
    (tmp_path / "delivery-receipts.json").write_text(json.dumps({"receipts": receipt}))
    assert delivery_status(tmp_path, output)["status"] == "approved"
    output.write_bytes(b"changed")
    assert delivery_status(tmp_path, output)["status"] == "blocked"
