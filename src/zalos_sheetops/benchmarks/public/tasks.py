import random
from collections import Counter

from .rubric import ORACLE_ISSUES


def select_tasks(
    tasks: list[dict], count: int = 30, seed: int = 20261009
) -> list[dict]:
    if count < 1 or count > len(tasks):
        raise ValueError("count must be between one and the dataset size")
    groups: dict[str, list[dict]] = {}
    for task in sorted(tasks, key=lambda t: str(t["id"])):
        groups.setdefault(task.get("instruction_type", "unknown"), []).append(task)
    rng = random.Random(seed)
    for group in groups.values():
        rng.shuffle(group)
    selected = []
    while len(selected) < count:
        for name in sorted(groups):
            if groups[name] and len(selected) < count:
                selected.append(groups[name].pop())
    return selected


def reviewed_selection(
    pool: list[dict], selected: list[dict], seed: int = 20261009
) -> list[dict]:
    """Keep every initial attempt; replace reviewed oracle gaps by type and seed."""
    result = selected.copy()
    identities = {str(t["id"]) for t in selected}
    wanted = Counter(t["instruction_type"] for t in selected)
    usable = Counter(
        t["instruction_type"] for t in selected if str(t["id"]) not in ORACLE_ISSUES
    )
    for task in select_tasks(pool, len(pool), seed):
        key, category = str(task["id"]), task["instruction_type"]
        if (
            key in identities
            or key in ORACLE_ISSUES
            or usable[category] >= wanted[category]
        ):
            continue
        result.append(task)
        identities.add(key)
        usable[category] += 1
    if any(usable[k] < n for k, n in wanted.items()):
        raise ValueError("Dataset cannot replace reviewed oracle gaps")
    return result
