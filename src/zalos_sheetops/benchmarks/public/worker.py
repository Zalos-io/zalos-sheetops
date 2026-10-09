"""One fresh process per measurement; stdout is only the result record."""

import json
import resource
import sys
import time

from .adapters import read_case
from .scoring import digest_rows


def main() -> None:
    case = json.loads(sys.argv[1])
    adapter = sys.argv[2]
    start = time.perf_counter()
    try:
        rows = read_case(case, adapter)
        result = {
            "status": "ok",
            "digest": digest_rows(rows),
            "rows": len(rows),
            "sample": rows[:20],
        }
    except NotImplementedError as error:
        result = {"status": "unsupported", "reason": str(error)}
    except Exception as error:
        result = {"status": "error", "reason": f"{type(error).__name__}: {error}"}
    usage = resource.getrusage(resource.RUSAGE_SELF)
    children = resource.getrusage(resource.RUSAGE_CHILDREN)
    result["operationSeconds"] = time.perf_counter() - start
    result["peakProcessRssBytes"] = max(usage.ru_maxrss, children.ru_maxrss) * (
        1 if sys.platform == "darwin" else 1024
    )
    result["cpuSeconds"] = (
        usage.ru_utime + usage.ru_stime + children.ru_utime + children.ru_stime
    )
    print(json.dumps(result, ensure_ascii=False, allow_nan=False))


if __name__ == "__main__":
    main()
