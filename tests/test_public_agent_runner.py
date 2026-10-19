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


def test_agent_process_starts_in_the_given_directory(tmp_path: Path) -> None:
    import os
    import sys

    from zalos_sheetops.benchmarks.public.agents import call_agent

    launch = tmp_path / "launch"
    launch.mkdir()
    command = [
        sys.executable,
        "-c",
        "import json, os; print(json.dumps([os.getcwd(), os.listdir()]))",
    ]
    measured = call_agent(command, dict(os.environ), 30, cwd=launch)
    cwd, entries = json.loads(measured["response"])
    assert Path(cwd).resolve() == launch.resolve()
    assert entries == []


def _run_one_task(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> list[dict]:
    """Run both arms of one task from a checkout, with relative paths."""
    import os

    from zalos_sheetops.benchmarks.public import agents

    folder = tmp_path / "data" / "spreadsheet" / "1"
    folder.mkdir(parents=True)
    (folder / "1_1_init.xlsx").write_bytes(b"input")
    (folder / "1_1_golden.xlsx").write_bytes(b"golden")
    for name in ["hermes", "SKILL.md", "soffice"]:
        (tmp_path / name).write_text(name)
    (tmp_path / "env").write_text("AI_GATEWAY_API_KEY=test\n")
    # The runner starts inside a checkout whose AGENTS.md is not part of the task.
    (tmp_path / "AGENTS.md").write_text("Repository instructions")
    monkeypatch.chdir(tmp_path)
    launches = []

    def fake_agent(
        command: list[str],
        env: dict[str, str],
        timeout: float,
        cwd: Path | None = None,
    ) -> dict:
        entries = sorted(os.listdir(cwd)) if cwd else None
        launches.append(
            {"cwd": cwd, "entries": entries, "command": command, "env": env}
        )
        return {
            "executionStatus": "completed",
            "seconds": 0.0,
            "exitCode": 0,
            "response": "",
            "stderrPresent": False,
        }

    monkeypatch.setattr(agents, "call_agent", fake_agent)
    monkeypatch.setattr(agents, "cleanup_workspace", lambda workspace: [])
    monkeypatch.setattr(agents, "grade_workbook", lambda *args: {"status": "fail"})
    task = {
        "id": 1,
        "spreadsheet_path": "spreadsheet/1",
        "instruction": "Fill B1.",
        "instruction_type": "Cell-Level Manipulation",
    }
    agents.run_tasks(
        Path("data"),
        [task],
        Path("out"),
        {"baseline": "baseline-image", "sheetops": "sheetops-image"},
        Path("hermes"),
        Path("env"),
        Path("SKILL.md"),
        Path("SKILL.md"),
        Path("soffice"),
        enforce_delivery=True,
    )
    return launches


def test_runner_launches_each_attempt_from_a_fresh_empty_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    launches = _run_one_task(tmp_path, monkeypatch)
    assert len(launches) == 2
    for launch in launches:
        assert launch["cwd"] is not None, "agent inherited the runner working directory"
        assert launch["entries"] == []
        assert not launch["cwd"].resolve().is_relative_to(tmp_path.resolve())
    assert launches[0]["cwd"] != launches[1]["cwd"]


def test_runner_gives_the_agent_absolute_paths(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    launches = _run_one_task(tmp_path, monkeypatch)
    for launch in launches:
        command, env = launch["command"], launch["env"]
        usage = Path(command[command.index("--usage-file") + 1])
        home = Path(env["HERMES_HOME"])
        config = json.loads((home / "config.yaml").read_text())
        volume = config["terminal"]["docker_volumes"][0].removesuffix(":/workspace")
        plugin = config["plugins"]["entries"]["sheetops-delivery"]["settings"]
        paths = [usage, home, Path(env["HERMES_RUNTIME_DIR"]), Path(volume)]
        assert all(path.is_absolute() for path in paths), paths
        assert Path(command[0]).resolve() == (tmp_path / "hermes").resolve()
        engine = Path(plugin["calculation_engine"]).resolve()
        assert engine == (tmp_path / "soffice").resolve()
