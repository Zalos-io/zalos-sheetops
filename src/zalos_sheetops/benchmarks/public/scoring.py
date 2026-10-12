import hashlib
import json
import statistics
from datetime import date, datetime, time
from decimal import Decimal


def scalar(value: object) -> object:
    if isinstance(value, date) and not isinstance(value, datetime):
        return datetime.combine(value, time()).isoformat()
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return value


def canonical(value: object) -> object:
    if type(value) in {int, float}:
        return {"number": str(Decimal(str(value)).normalize())}
    if isinstance(value, list):
        return [canonical(item) for item in value]
    return scalar(value)


def digest_rows(rows: list[list[object]]) -> str:
    digest = hashlib.sha256()
    for row in rows:
        digest.update(
            json.dumps(
                canonical(row),
                ensure_ascii=False,
                allow_nan=False,
                separators=(",", ":"),
            ).encode("utf-8")
        )
        digest.update(b"\n")
    return digest.hexdigest()


def score_outputs(expected: list, actual: list) -> dict:
    differences = []
    count = 0
    for row in range(max(len(expected), len(actual))):
        left = expected[row] if row < len(expected) else []
        right = actual[row] if row < len(actual) else []
        for col in range(max(len(left), len(right))):
            a = left[col] if col < len(left) else {"missing": True}
            b = right[col] if col < len(right) else {"missing": True}
            if canonical(a) != canonical(b):
                count += 1
                if len(differences) < 10:
                    differences.append(
                        {"row": row + 1, "column": col + 1, "expected": a, "actual": b}
                    )
    return {
        "correct": count == 0,
        "mismatchCount": count,
        "firstDifferences": differences,
    }


def summarize(records: list[dict]) -> list[dict]:
    groups: dict[tuple[str, str], list[dict]] = {}
    for record in records:
        groups.setdefault((record["case"], record["adapter"]), []).append(record)
    result = []
    for (case, adapter), runs in sorted(groups.items()):
        passes = [r for r in runs if r["status"] == "pass"]
        complete = len(passes) == len(runs)
        times = [r["seconds"] for r in passes]
        result.append(
            {
                "case": case,
                "adapter": adapter,
                "attempts": len(runs),
                "passes": len(passes),
                "passRate": len(passes) / len(runs),
                "medianSeconds": statistics.median(times) if complete else None,
                "minSeconds": min(times) if complete else None,
                "maxSeconds": max(times) if complete else None,
            }
        )
    return result
