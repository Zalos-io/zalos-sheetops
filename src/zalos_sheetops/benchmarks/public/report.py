"""Aggregate correctness before speed; never discard failed or unscorable attempts."""

import argparse
import json
import re
import statistics
from collections import Counter
from pathlib import Path

from .scoring import summarize

UNSCORABLE = {
    "oracle_cache_missing",
    "output_cache_missing",
    "ambiguous_task_metadata",
    "oracle_failure",
    "oracle_preservation_conflict",
    "oracle_style_conflict",
    "recalculation_failure",
    "source_mutated",
    "evaluation_unavailable",
}


def read_records(path: Path) -> list[dict]:
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text().splitlines() if line]


def coverage(records: list[dict], expected: set[tuple], keys: tuple[str, ...]) -> dict:
    observed = Counter(tuple(str(r[k]) for k in keys) for r in records)
    wanted = {tuple(str(value) for value in key) for key in expected}
    return {
        "complete": set(observed) == wanted and all(n == 1 for n in observed.values()),
        "missing": len(wanted - set(observed)),
        "unexpected": len(set(observed) - wanted),
        "duplicates": sum(n - 1 for n in observed.values()),
    }


def adoption(runs: list[dict]) -> dict:
    commands, errors = Counter(), Counter()
    for run in runs:
        commands.update(run["usage"].get("sheetopsCommandCounts", {}))
        errors.update(run["usage"].get("sheetopsErrorCounts", {}))
    return {
        "tasksWithResponses": sum(
            bool(r["usage"].get("sheetopsCommandCounts")) for r in runs
        ),
        "successfulResponses": sum(
            r["usage"].get("sheetopsSuccessfulResponses", 0) for r in runs
        ),
        "failedResponses": sum(
            r["usage"].get("sheetopsFailedResponses", 0) for r in runs
        ),
        "commandCounts": dict(commands),
        "errorCounts": dict(errors),
        "scope": "observed JSON command responses; Python API use is not counted",
    }


def operation_table(records: list[dict]) -> list[str]:
    adapters = sorted({r["adapter"] for r in records})
    groups = {}
    for record in records:
        groups.setdefault((record["case"], record["adapter"]), []).append(record)
    summaries = {(r["case"], r["adapter"]): r for r in summarize(records)}
    lines = [
        "| Case | " + " | ".join(adapters) + " |",
        "|---|" + "---:|" * len(adapters),
    ]
    for case in sorted({r["case"] for r in records}):
        cells = []
        for adapter in adapters:
            runs = groups.get((case, adapter), [])
            summary = summaries.get((case, adapter), {})
            seconds = summary.get("medianSeconds")
            if seconds is None:
                counts = Counter(r["status"] for r in runs)
                cells.append(
                    ", ".join(f"{k} ({v})" for k, v in sorted(counts.items()))
                    or "pending"
                )
            else:
                rss = [
                    r["peakProcessRssBytes"] / 1048576
                    for r in runs
                    if "peakProcessRssBytes" in r
                ]
                memory = f" / {statistics.median(rss):.1f} MiB" if rss else ""
                cells.append(f"{seconds:.3f}s{memory}")
        lines.append(f"| {case} | " + " | ".join(cells) + " |")
    return lines


def summarize_agents(records: list[dict]) -> dict:
    arms = {}
    pairs = {}
    for record in records:
        pairs.setdefault((record["task"], record["repeat"]), {})[record["arm"]] = record
    for arm in ["baseline", "sheetops"]:
        runs = [r for r in records if r["arm"] == arm]
        costs_available = bool(runs) and all(
            r["usage"].get("cost_status") not in {None, "unknown"}
            and isinstance(r["usage"].get("estimated_cost_usd"), (int, float))
            for r in runs
        )
        arms[arm] = {
            "attempts": len(runs),
            "scorableAttempts": sum(
                r["grading"]["status"] not in UNSCORABLE for r in runs
            ),
            "unscorableAttempts": sum(
                r["grading"]["status"] in UNSCORABLE for r in runs
            ),
            "statusCounts": dict(Counter(r["grading"]["status"] for r in runs)),
            "totalSecondsAllAttempts": sum(r["seconds"] for r in runs),
            "totalMainTokens": sum(r["usage"].get("total_tokens", 0) for r in runs),
            "totalObservedToolCalls": sum(
                r["usage"].get("observedToolCalls", 0) for r in runs
            ),
            "sheetopsAdoption": adoption(runs),
            "costAvailable": costs_available,
            "estimatedCostUsd": sum(r["usage"]["estimated_cost_usd"] for r in runs)
            if costs_available
            else None,
        }
    complete = [p for p in pairs.values() if set(p) == {"baseline", "sheetops"}]
    correct = [
        p for p in complete if all(r["grading"]["status"] == "pass" for r in p.values())
    ]
    scorable = [
        p
        for p in complete
        if all(r["grading"]["status"] not in UNSCORABLE for r in p.values())
    ]
    ratios = [p["baseline"]["seconds"] / p["sheetops"]["seconds"] for p in correct]
    return {
        "attemptCount": len(records),
        "completePairCount": len(complete),
        "scorablePairCount": len(scorable),
        "correctPairCount": len(correct),
        "arms": arms,
        "medianPairedSpeedup": statistics.median(ratios) if ratios else None,
        "speedComparisonScope": (
            "only pairs where both arms pass; selection bias remains"
        ),
        "attribution": (
            "package plus skill configuration; not package-only or Electron UI"
        ),
    }


def has_date_gap(record: dict) -> bool:
    for difference in record.get("sampleComparison", {}).get("firstDifferences", []):
        wanted, actual = difference.get("expected"), difference.get("actual")
        if isinstance(wanted, list) and isinstance(actual, list):
            wanted, actual = wanted[0], actual[0]
        if isinstance(wanted, str) and re.match(r"\d{4}-\d{2}-\d{2}[T ]", wanted):
            if re.fullmatch(r"\d+(?:\.\d+)?", str(actual)):
                return True
    return False


def failure_catalog(operations: list[dict], agents: list[dict]) -> list[dict]:
    catalog = []
    tool_errors = [
        r for r in agents if r.get("usage", {}).get("sheetopsFailedResponses")
    ]
    if tool_errors:
        errors = Counter()
        for record in tool_errors:
            errors.update(record["usage"].get("sheetopsErrorCounts", {}))
        catalog.append(
            {
                "priority": 3,
                "id": "tool-contract-misuse",
                "kind": "observed tool errors; task may still pass",
                "title": "Reduce invalid SheetOps calls in the spreadsheet skill",
                "failedAttempts": len(tool_errors),
                "errorCounts": dict(errors),
                "taskIds": sorted({str(r["task"]) for r in tool_errors}),
                "evidence": "Command JSON returned a stable error; inspect the "
                "task transcript before attributing a package defect",
                "examples": [
                    {
                        "task": str(r["task"]),
                        "arm": r["arm"],
                        "repeat": r["repeat"],
                        "errors": r["usage"]["sheetopsErrorCounts"],
                    }
                    for r in tool_errors[:3]
                ],
            }
        )
    failed_ops = [
        r
        for r in operations
        if r["adapter"].startswith("sheetops") and r["status"] != "pass"
    ]
    date_failures = [r for r in failed_ops if has_date_gap(r)]
    other_failures = [r for r in failed_ops if not has_date_gap(r)]
    if date_failures:
        catalog.append(
            {
                "priority": 1,
                "id": "typed-date-extraction",
                "kind": "capability gap pending contract review",
                "title": "Expose date-aware values in bounded XML and table reads",
                "evidence": (
                    "Date-formatted cells return stored Excel serials; typed reader "
                    "oracle returns ISO dates"
                ),
                "failedAttempts": len(date_failures),
                "caseIds": sorted({r["case"] for r in date_failures}),
                "reproduce": (
                    "uv run python -m zalos_sheetops.benchmarks.public.runner replay "
                    "--data DATA --output REPLAY --case "
                    + sorted({r["case"] for r in date_failures})[0]
                    + " --adapter sheetops-cli"
                ),
                "examples": [
                    {
                        "case": r["case"],
                        "adapter": r["adapter"],
                        "differences": r.get("sampleComparison", {}).get(
                            "firstDifferences", []
                        )[:2],
                    }
                    for r in date_failures[:2]
                ],
            }
        )
    if other_failures:
        catalog.append(
            {
                "priority": 1,
                "id": "sheetops-operation-contract",
                "kind": "reader contract failure",
                "title": "Investigate other reader contract failures",
                "failedAttempts": len(other_failures),
                "caseIds": sorted({r["case"] for r in other_failures}),
            }
        )
    groups = {}
    for record in agents:
        grading = record["grading"]
        status = grading["status"]
        if status == "pass":
            continue
        if status in UNSCORABLE:
            kind = "oracle-coverage"
        elif status != "incorrect":
            kind = status
        else:
            counts = grading.get("failureCounts", {})
            categories = {
                "values": "answer-values",
                "styles": "explicit-styles",
                "formulas": "formula-requirement",
                "preservation": "workbook-preservation",
                "structure": "workbook-structure",
            }
            kinds = [
                name for component, name in categories.items() if counts.get(component)
            ]
            formula_errors = [
                d
                for d in grading.get("valueDifferences", [])
                if isinstance(d.get("actual"), str) and d["actual"].startswith("#")
            ]
            if formula_errors:
                kinds.append("formula-evaluation")
                if counts.get("values", 0) == len(formula_errors):
                    kinds = [k for k in kinds if k != "answer-values"]
            if not kinds:
                kinds = (
                    ["workbook-preservation"]
                    if grading.get("preservationFailures")
                    or grading.get("structureDifferences")
                    else ["golden-values-or-styles"]
                )
            for kind in kinds:
                groups.setdefault(kind, []).append(record)
            continue
        groups.setdefault(kind, []).append(record)
    for kind, records in sorted(groups.items()):
        titles = {
            "answer-values": "Validate answer values and string boundaries",
            "explicit-styles": "Apply the requested visible formatting",
            "workbook-preservation": "Restrict writes and preserve unrelated cells",
            "workbook-structure": "Preserve workbook structure",
            "invalid_output": "Validate XLSX formula serialization in native Excel",
            "formula-evaluation": "Validate formulas with the target Excel engine",
            "timeout": "Bound workbook work and surface incomplete results",
            "oracle-coverage": "Resolve reference conflicts before scoring",
        }
        catalog.append(
            {
                "priority": 1
                if kind in {"workbook-preservation", "invalid_output"}
                else 3
                if kind == "oracle-coverage"
                else 2,
                "id": kind,
                "kind": "agent task failure"
                if kind != "oracle-coverage"
                else "unscorable dataset/evaluator case",
                "title": titles.get(kind, kind.replace("-", " ")),
                "failedAttempts": len(records),
                "taskIds": sorted({str(r["task"]) for r in records}),
                "arms": dict(Counter(r["arm"] for r in records)),
                "examples": [
                    {
                        "task": str(r["task"]),
                        "arm": r["arm"],
                        "repeat": r["repeat"],
                        "grading": r["grading"],
                    }
                    for r in records[:3]
                ],
                "reproduce": (
                    "Run agents.py with the recorded task, model, skill hashes and "
                    "images"
                ),
                "scope": "Formula errors use each record's evaluation engine; "
                "native Excel and LibreOffice outcomes are reported separately"
                if kind == "formula-evaluation"
                else "recorded task rubric",
            }
        )
    return sorted(catalog, key=lambda r: (r["priority"], -r["failedAttempts"], r["id"]))


def build_report(root: Path, output: Path, agent_records: Path | None = None) -> dict:
    operations = read_records(root / "baseline/operations.jsonl")
    contracts = read_records(root / "contracts/contracts.jsonl")
    agents = read_records(agent_records or root / "agents/agents.jsonl")
    operation_metadata_path = root / "baseline/metadata.json"
    operation_metadata = (
        json.loads(operation_metadata_path.read_text())
        if operation_metadata_path.exists()
        else {}
    )
    adapters = operation_metadata.get(
        "adapters",
        ["openpyxl", "sheetops-api", "sheetops-cli", "calamine", "fastexcel"],
    )
    expected_operations = (
        len(operation_metadata.get("cases", []))
        * operation_metadata.get("repeats", 1)
        * len(adapters)
    )
    metadata_path = root / "agents/agent-metadata.json"
    metadata = json.loads(metadata_path.read_text()) if metadata_path.exists() else {}
    expected_pairs = len(metadata.get("taskIds", [])) * metadata.get("repeats", 1)
    operation_coverage = coverage(
        operations,
        {
            (case["id"], adapter, repeat)
            for case in operation_metadata.get("cases", [])
            for adapter in adapters
            for repeat in range(operation_metadata.get("repeats", 1))
        },
        ("case", "adapter", "repeat"),
    )
    agent_coverage = coverage(
        agents,
        {
            (task, arm, repeat)
            for task in metadata.get("taskIds", [])
            for arm in ["baseline", "sheetops"]
            for repeat in range(metadata.get("repeats", 1))
        },
        ("task", "arm", "repeat"),
    )
    result = {
        "schemaVersion": 1,
        "status": "complete"
        if (
            len(agents) == 2 * expected_pairs
            and expected_pairs
            and expected_operations
            and len(operations) == expected_operations
            and contracts
            and operation_coverage["complete"]
            and agent_coverage["complete"]
        )
        else "in_progress",
        "operationMeasurements": len(operations),
        "coverage": {"operations": operation_coverage, "agents": agent_coverage},
        "expectedOperationMeasurements": expected_operations,
        "agentRecordSource": str(agent_records or root / "agents/agents.jsonl"),
        "operationSummaries": summarize(operations),
        "operationStatusCounts": {
            adapter: dict(
                Counter(r["status"] for r in operations if r["adapter"] == adapter)
            )
            for adapter in sorted({r["adapter"] for r in operations})
        },
        "featureContracts": {
            "attempts": len(contracts),
            "statusCounts": dict(Counter(r["status"] for r in contracts)),
            "fixtureManifestDifferences": sum(
                "fixtureManifestDifference" in r for r in contracts
            ),
        },
        "agents": summarize_agents(agents),
        "expectedAgentPairs": expected_pairs,
        "evaluationEngines": sorted(
            {
                json.dumps(r["evaluationEngine"], sort_keys=True)
                for r in agents
                if "evaluationEngine" in r
            }
        ),
        "rankedFailures": failure_catalog(operations, agents),
        "limitations": [
            "Failure groups overlap: one attempt can fail values, styles "
            "and preservation",
            (
                "Fresh processes, uncontrolled OS file cache; one host; three "
                "operation repetitions"
            ),
            "RSS is maximum for one worker or CLI subprocess, not aggregate memory",
            (
                "Fidelity-prefixed operation cases compare cell "
                "values/formulas/caches, not every workbook feature"
            ),
            (
                "Feature contracts separately cover inventory, merges, number "
                "formats, hyperlinks and decimal reconciliation"
            ),
            (
                "Agent scores use golden answer-region values, explicit styles "
                "and preservation checks; not official leaderboard scores"
            ),
            "Formula evaluation engines are recorded per attempt; native Excel "
            "caches are audited independently of raw outputs. LibreOffice results "
            "are a separate interoperability check when native results are available",
            "Native Excel opens formula-containing evaluation copies; value-only "
            "outputs are parsed without a native UI open check",
            "Observed SheetOps commands are inspection and reads; workbook edits "
            "use agent-authored Python. Output failures do not establish a "
            "SheetOps write API defect",
            "Task scores verify the supplied workbook result; future-sheet macro "
            "behavior and VBA execution are not tested",
            "The first 24 agent attempts lack per-attempt execution source hashes; "
            "image, launcher, skill and input hashes remain recorded",
            (
                "One agent repetition per arm in this initial baseline; "
                "insufficient for a stable end-to-end speed claim"
            ),
            (
                "Inference cost is unavailable when provider cost status is "
                "unknown, even if reported number is zero"
            ),
        ],
    }
    output.mkdir(parents=True, exist_ok=True)
    (output / "report.json").write_text(json.dumps(result, indent=2))
    lines = [
        "# Public spreadsheet benchmark baseline",
        "",
        f"Run status: {result['status']}.",
        "",
        f"Operations: {len(operations)} measurements. "
        f"Feature contracts: {len(contracts)} checks.",
        "",
        "| Reader | Pass | Incorrect | Unsupported/error |",
        "|---|---:|---:|---:|",
    ]
    for adapter, counts in result["operationStatusCounts"].items():
        lines.append(
            f"| {adapter} | {counts.get('pass', 0)} | "
            f"{counts.get('incorrect', 0)} | "
            f"{sum(v for k, v in counts.items() if k not in {'pass', 'incorrect'})} |"
        )
    lines.extend(
        [
            "",
            "Reader timing and memory",
            "",
            "Cells show median wall time including worker startup / "
            "median peak single-process RSS. "
            "Timings appear only when every recorded repetition passes. "
            "Date-containing ranges use a typed-value contract; "
            "scalar ranges exclude the date column.",
            "",
            *operation_table(operations),
            "",
            "Paired agent tasks",
            "",
            f"Completed pairs: {result['agents']['completePairCount']}"
            f"/{expected_pairs}; "
            f"scorable pairs: {result['agents']['scorablePairCount']}; "
            f"both-correct pairs: {result['agents']['correctPairCount']}.",
            "Evaluation engines: "
            f"{', '.join(result['evaluationEngines']) or 'not recorded'}.",
            "",
        ]
    )
    for arm, summary in result["agents"]["arms"].items():
        lines.append(
            f"- {arm}: {summary['attempts']} attempts; "
            f"{summary['scorableAttempts']} scorable / "
            f"{summary['unscorableAttempts']} unscorable; "
            f"outcomes {summary['statusCounts']}; inference cost "
            f"{'available' if summary['costAvailable'] else 'unavailable'}."
        )
        usage = summary["sheetopsAdoption"]
        lines.append(
            f"  Observed SheetOps JSON responses: {usage['tasksWithResponses']} tasks, "
            f"{usage['successfulResponses']} successful / "
            f"{usage['failedResponses']} failed; "
            f"commands {usage['commandCounts']}; errors {usage['errorCounts']}."
        )
    lines.extend(["", "Ranked follow-ups", ""])
    for failure in result["rankedFailures"]:
        ids = failure.get("taskIds", failure.get("caseIds", []))
        lines.append(
            f"- P{failure['priority']} {failure['title']}: "
            f"{failure['failedAttempts']} affected attempts. "
            f"Reproduction IDs: {', '.join(ids)}."
        )
        if failure.get("reproduce"):
            lines.extend(["", f"  Reproduce: `{failure['reproduce']}`", ""])
    lines.extend(
        [
            "",
            "Measurement limits",
            "",
            *[f"- {item}" for item in result["limitations"]],
            "",
            "Source attribution",
            "",
            "- Mixed-column generator adapted from nazmul284/excel-read-benchmark "
            "(MIT), revision d8ddc2b5cfab51d53e196015d7037b1d5f463a8a. "
            "https://github.com/nazmul284/excel-read-benchmark/tree/d8ddc2b5cfab51d53e196015d7037b1d5f463a8a",
            "- Selected feature workbooks from SynthGL/ExcelBench (MIT), revision "
            "5f71713413c43f3733e303ba8a4eed6a756760cd. "
            "https://github.com/SynthGL/ExcelBench/tree/5f71713413c43f3733e303ba8a4eed6a756760cd",
            "- Selected Verified-400 tasks from KAKA22/SpreadsheetBench "
            "(CC-BY-SA-4.0), "
            "revision ab0b742b0fc95b946f212d80ac7771b5531272e4. "
            "https://huggingface.co/datasets/KAKA22/SpreadsheetBench/tree/ab0b742b0fc95b946f212d80ac7771b5531272e4",
            "",
            "These are our measurements and scoped checks, not the upstream "
            "projects' published results or leaderboard scores. "
            "Source URLs and content hashes are recorded in sources.json; "
            "runtime versions, input hashes and code hashes are in run metadata.",
            "",
        ]
    )
    (output / "baseline.md").write_text("\n".join(lines))
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--agent-records", type=Path)
    args = parser.parse_args()
    result = build_report(args.root, args.output, args.agent_records)
    print(
        result["status"],
        result["operationMeasurements"],
        result["agents"]["completePairCount"],
    )


if __name__ == "__main__":
    main()
