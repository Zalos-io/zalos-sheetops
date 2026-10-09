"""Aggregate correctness before speed; never discard failed or unscorable attempts."""

import argparse
import json
import re
import statistics
from collections import Counter
from pathlib import Path

from .scoring import summarize


def read_records(path: Path) -> list[dict]:
    if not path.exists():
        return []
    return [json.loads(line) for line in path.read_text().splitlines() if line]


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
            "statusCounts": dict(Counter(r["grading"]["status"] for r in runs)),
            "totalSecondsAllAttempts": sum(r["seconds"] for r in runs),
            "totalMainTokens": sum(r["usage"].get("total_tokens", 0) for r in runs),
            "totalObservedToolCalls": sum(
                r["usage"].get("observedToolCalls", 0) for r in runs
            ),
            "costAvailable": costs_available,
            "estimatedCostUsd": sum(r["usage"]["estimated_cost_usd"] for r in runs)
            if costs_available
            else None,
        }
    complete = [p for p in pairs.values() if set(p) == {"baseline", "sheetops"}]
    correct = [
        p for p in complete if all(r["grading"]["status"] == "pass" for r in p.values())
    ]
    ratios = [p["baseline"]["seconds"] / p["sheetops"]["seconds"] for p in correct]
    return {
        "attemptCount": len(records),
        "completePairCount": len(complete),
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
                    "Use recorded metadata case with the public worker for "
                    "sheetops-api and openpyxl"
                ),
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
        if status in {
            "oracle_cache_missing",
            "ambiguous_task_metadata",
            "oracle_failure",
            "recalculation_failure",
        }:
            kind = "oracle-coverage"
        elif status != "incorrect":
            kind = status
        elif grading.get("preservationFailures") or grading.get("structureDifferences"):
            kind = "workbook-preservation"
        else:
            kind = "golden-values-or-styles"
        groups.setdefault(kind, []).append(record)
    for kind, records in sorted(groups.items()):
        catalog.append(
            {
                "priority": 2 if kind != "oracle-coverage" else 3,
                "id": kind,
                "kind": "agent task failure"
                if kind != "oracle-coverage"
                else "unscorable dataset/evaluator case",
                "title": kind.replace("-", " "),
                "failedAttempts": len(records),
                "taskIds": sorted({str(r["task"]) for r in records}),
                "arms": dict(Counter(r["arm"] for r in records)),
                "reproduce": (
                    "Run agents.py with the recorded task, model, skill hashes and "
                    "images"
                ),
            }
        )
    return sorted(catalog, key=lambda r: (r["priority"], -r["failedAttempts"], r["id"]))


def build_report(root: Path, output: Path) -> dict:
    operations = read_records(root / "baseline/operations.jsonl")
    contracts = read_records(root / "contracts/contracts.jsonl")
    agents = read_records(root / "agents/agents.jsonl")
    metadata_path = root / "agents/agent-metadata.json"
    metadata = json.loads(metadata_path.read_text()) if metadata_path.exists() else {}
    expected_pairs = len(metadata.get("taskIds", [])) * metadata.get("repeats", 1)
    result = {
        "schemaVersion": 1,
        "status": "complete"
        if len(agents) == 2 * expected_pairs and expected_pairs
        else "in_progress",
        "operationMeasurements": len(operations),
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
        "rankedFailures": failure_catalog(operations, agents),
        "limitations": [
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
                "Agent task scores use strict whole-workbook golden "
                "values/styles plus unchanged-cell preservation; not official "
                "leaderboard scores"
            ),
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
            "Paired agent tasks",
            "",
            f"Completed pairs: {result['agents']['completePairCount']}"
            f"/{expected_pairs}; "
            f"both-correct pairs: {result['agents']['correctPairCount']}.",
            "",
        ]
    )
    for arm, summary in result["agents"]["arms"].items():
        lines.append(
            f"- {arm}: {summary['attempts']} attempts; "
            f"outcomes {summary['statusCounts']}; inference cost "
            f"{'available' if summary['costAvailable'] else 'unavailable'}."
        )
    lines.extend(["", "Ranked follow-ups", ""])
    for failure in result["rankedFailures"]:
        ids = failure.get("taskIds", failure.get("caseIds", []))
        lines.append(
            f"- P{failure['priority']} {failure['title']}: "
            f"{failure['failedAttempts']} failed/unscorable attempts. "
            f"Reproduction IDs: {', '.join(ids)}."
        )
    lines.extend(
        [
            "",
            "Measurement limits",
            "",
            *[f"- {item}" for item in result["limitations"]],
            "",
        ]
    )
    (output / "baseline.md").write_text("\n".join(lines))
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = build_report(args.root, args.output)
    print(
        result["status"],
        result["operationMeasurements"],
        result["agents"]["completePairCount"],
    )


if __name__ == "__main__":
    main()
