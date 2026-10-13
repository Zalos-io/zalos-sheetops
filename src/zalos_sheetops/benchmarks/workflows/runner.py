"""Three-arm agent comparison with isolated inputs and held-out oracles."""

import argparse
import json
import os
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from zalos_sheetops.benchmarks.public.agents import (
    call_agent,
    cleanup_workspace,
    containers_for_workspace,
    extract_usage,
    provider_env,
)
from zalos_sheetops.benchmarks.public.fixtures import file_sha256

from .cases import build_cases
from .grade import grade_case


def probe_images(images: dict[str, str]) -> dict:
    probes = {}
    script = (
        "import json, shutil, subprocess, importlib.metadata as m; "
        "engine=shutil.which('soffice') or shutil.which('libreoffice'); "
        "print(json.dumps({'engine': engine, "
        "'engineVersion': subprocess.check_output([engine, '--version'], "
        "text=True).strip() if engine else None, "
        "'libraries': {n:m.version(n) for n in "
        "['openpyxl','polars','fastexcel','pyarrow','xlsxwriter']}}))"
    )
    for image in dict.fromkeys(images.values()):
        probe = json.loads(
            subprocess.check_output(
                [
                    "docker",
                    "run",
                    "--rm",
                    "--network",
                    "none",
                    "--entrypoint",
                    "python",
                    image,
                    "-c",
                    script,
                ],
                text=True,
                timeout=60,
            )
        )
        if not probe.get("engine"):
            raise ValueError(f"Benchmark image lacks a calculation engine: {image}")
        probes[image] = probe
    return probes


def container_peak(workspace: Path) -> int | None:
    ids = subprocess.check_output(
        ["docker", "ps", "-aq", "--filter", "label=hermes-agent=1"], text=True
    ).split()
    if not ids:
        return None
    containers = json.loads(
        subprocess.check_output(["docker", "inspect", *ids], text=True)
    )
    peaks = []
    for identity in containers_for_workspace(containers, workspace):
        result = subprocess.run(
            ["docker", "exec", identity, "cat", "/sys/fs/cgroup/memory.peak"],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0 and result.stdout.strip().isdigit():
            peaks.append(int(result.stdout))
    return max(peaks) if peaks else None


def run_one(case: dict, arm: str, repeat: int, settings: dict, output: Path) -> dict:
    root = output / "runs" / case["id"] / arm / str(repeat)
    if root.exists():
        raise ValueError(f"Run exists; use a new output directory: {root}")
    workspace, home = root / "workspace", root / "home"
    shutil.copytree(case["inputs"], workspace)
    home.mkdir()
    input_hashes = {p.name: file_sha256(p) for p in workspace.iterdir()}
    config = {
        "model": {"provider": "ai-gateway", "default": settings["model"]},
        "agent": {"max_turns": 30},
        "memory": {"memory_enabled": False, "user_profile_enabled": False},
        "terminal": {
            "backend": "docker",
            "cwd": "/workspace",
            "docker_image": settings["images"][arm],
            "docker_image_pinned": True,
            "docker_volumes": [f"{workspace}:/workspace"],
            "docker_mount_cwd_to_workspace": False,
            "docker_forward_env": [],
            "docker_env": {},
            "docker_network": False,
            "docker_orphan_reaper": False,
            "docker_persist_across_processes": True,
            "container_persistent": True,
            "container_cpu": 1,
            "container_memory": 2048,
        },
        "mcp_servers": {},
        "fallback_models": [],
        "fallback_providers": [],
        "display": {"resume_session": False},
        "plugins": {"enabled": []},
    }
    (home / "config.yaml").write_text(json.dumps(config), encoding="utf-8")
    skill_path = Path(
        settings["baselineSkill"] if arm == "baseline" else settings["improvedSkill"]
    )
    prompt = (
        "Complete this spreadsheet task in /workspace. Input files are already there. "
        "Preserve the original inputs. Use the supplied spreadsheet skill. "
        "Helpers are at "
        "/opt/skills/xlsx/scripts and references at /opt/skills/xlsx/references. "
        "Network is unavailable. Do not inspect files outside this workspace and those "
        "skill directories. Save the requested output under /workspace.\n\nTask:\n"
        + case["instruction"]
        + "\n\nSkill:\n"
        + skill_path.read_text()
    )
    env = {
        **os.environ,
        **provider_env(Path(settings["envFile"])),
        "HERMES_HOME": str(home),
        "HERMES_PROFILE": "",
        "TERMINAL_CWD": "/workspace",
        "HERMES_RUNTIME_DIR": str(root / "runtime"),
    }
    for key in list(env):
        if key.startswith("ZALOS_") or key in {"HERMES_ENV_PATH", "HERMES_CONFIG_PATH"}:
            env.pop(key)
    usage = root / "usage.json"
    command = [
        settings["launcher"],
        "-z",
        prompt,
        "--model",
        settings["model"],
        "--provider",
        "ai-gateway",
        "--reasoning",
        "medium",
        "--toolsets",
        "terminal",
        "--usage-file",
        str(usage),
    ]
    try:
        measured = call_agent(command, env, settings["timeoutSeconds"])
        peak = container_peak(workspace)
    finally:
        cleanup_workspace(workspace)
    response = measured.pop("response")
    (root / "response.txt").write_text(response, encoding="utf-8")
    try:
        grade = grade_case(case, workspace, Path(settings["soffice"]))
    except Exception as error:
        grade = {
            "correct": False,
            "failures": [type(error).__name__ + ": " + str(error)],
            "formulaEvaluation": "error",
        }
    unchanged = all(
        (workspace / name).exists() and file_sha256(workspace / name) == digest
        for name, digest in input_hashes.items()
    )
    grade["inputsPreserved"] = unchanged
    grade["correct"] = (
        grade["correct"] and unchanged and measured["executionStatus"] == "completed"
    )
    record = {
        "case": case["id"],
        "arm": arm,
        "repeat": repeat,
        **measured,
        "peakContainerBytes": peak,
        "grading": grade,
        "usage": extract_usage(home, usage),
        "inputSha256": input_hashes,
        "outputSha256": file_sha256(workspace / case["output"])
        if (workspace / case["output"]).exists()
        else None,
    }
    (root / "result.json").write_text(json.dumps(record, indent=2), encoding="utf-8")
    print(
        case["id"],
        arm,
        repeat,
        grade["correct"],
        round(measured["seconds"], 2),
        flush=True,
    )
    return record


def run(
    settings: dict, output: Path, cases: list[dict], repeats: int, workers: int
) -> list[dict]:
    probes = probe_images(settings["images"])
    output.mkdir(parents=True, exist_ok=True)
    safe = {k: v for k, v in settings.items() if k != "envFile"}
    metadata = {
        **safe,
        "repeats": repeats,
        "workers": workers,
        "imageProbes": probes,
        "scope": "Desktop source agent core via CLI; not Electron UI E2E",
        "skillSha256": {
            name: file_sha256(Path(settings[name]))
            for name in ["baselineSkill", "improvedSkill"]
        },
        "caseIds": [c["id"] for c in cases],
        "deliveryEnforcement": False,
        "peakMemoryScope": (
            "terminal container cgroup peak including filesystem cache; "
            "null if unavailable"
        ),
        "sourceSha256": {
            p.name: file_sha256(p) for p in Path(__file__).parent.glob("*.py")
        },
    }
    (output / "metadata.json").write_text(
        json.dumps(metadata, indent=2), encoding="utf-8"
    )
    jobs = []
    arms = ["baseline", "skill_only", "sheetops"]
    for repeat in range(repeats):
        for index, case in enumerate(cases):
            offset = (repeat + index) % len(arms)
            for arm in arms[offset:] + arms[:offset]:
                jobs.append((case, arm, repeat))

    def execute(job: tuple) -> dict:
        return run_one(*job, settings, output)

    with (
        ThreadPoolExecutor(max_workers=workers) as executor,
        (output / "results.jsonl").open("w") as destination,
    ):
        records = []
        for result in executor.map(execute, jobs):
            records.append(result)
            destination.write(json.dumps(result) + "\n")
            destination.flush()
    return records


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--settings", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--fixtures", type=Path, required=True)
    parser.add_argument("--repeats", type=int, default=2)
    parser.add_argument("--workers", type=int, default=2)
    parser.add_argument("--case")
    args = parser.parse_args()
    manifest = args.fixtures / "cases.json"
    cases = (
        json.loads(manifest.read_text())
        if manifest.exists()
        else build_cases(args.fixtures)
    )
    if args.case:
        cases = [case for case in cases if case["id"] == args.case]
    run(
        json.loads(args.settings.read_text()),
        args.output,
        cases,
        args.repeats,
        args.workers,
    )


if __name__ == "__main__":
    main()
