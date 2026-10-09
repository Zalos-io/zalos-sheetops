import random


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
