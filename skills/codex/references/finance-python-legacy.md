# Low-level finance Python fallback

Use when `zalos_sheetops.finance` is unavailable or a lower-level operation is
needed. Preserve all complete-population and independent business controls.

## Financial controls

```python
from zalos_sheetops.reconcile.exact import reconcile_exact
from zalos_sheetops.validation.totals import check_totals

# Pass whole table dicts, retaining population metadata, not table["rows"].
# CSV already has dict rows; adapt each XLSX side first:
# left = {**left, "rows": [dict(zip(left["headers"], row)) for row in left["rows"]]}
comparison = reconcile_exact(left, right, ["reference"], ["reference"], "amount", "amount", "0")
control = check_totals(left_amounts, right_amounts, decimal_places=2, rounding="half_up")
```

`reconcile_exact` checks both whole sides including unmatched rows: duplicate/
blank keys, invalid amounts and incomplete metadata reject. Returns `matched`,
`leftUnmatched`, `rightUnmatched`, `variance` lists and `summary` counts. Variance:
`key: dict[column, str]`, `leftAmount`, `rightAmount`, exact `difference`;
`item["key"]["reference"]` is the text reference.

For balance/rounding controls, compare complete actual populations separately per
group/currency in the same units, using source decimal strings:

```python
from decimal import Context, Decimal, ROUND_HALF_UP, ROUND_HALF_EVEN, localcontext
import json
from pathlib import Path
from zalos_sheetops.io.read_csv_profile import read_csv_profile
from zalos_sheetops.validation.totals import check_totals
# Bind actual path, normalized column names, side labels and per-group precision.
group_column, side_column, amount_column = "currency", "side", "amount"
places_by_group = {"GBP": 2, "JPY": 0}  # replace with the task's complete mapping
left_label, right_label, rounding, source_tolerance = "left", "right", "half_up", "0"
table, profile = read_csv_profile(Path("source.csv"),
    key_columns=[group_column, side_column], amount_columns=[amount_column])
buckets = {group: {left_label: [], right_label: []} for group in places_by_group}
for row in table["rows"]:  # every validated record assigned exactly once
    assert row[group_column] in buckets and row[side_column] in {left_label, right_label}
    buckets[row[group_column]][row[side_column]].append(row[amount_column])
assert sum(len(values) for sides in buckets.values() for values in sides.values()) == profile["rowCount"]
groups = []
for group, places in sorted(places_by_group.items()):
    left, right = buckets[group][left_label], buckets[group][right_label]
    assert left and right
    report = check_totals(left, right, decimal_places=places, rounding=rounding,
                          source_tolerance=source_tolerance)
    assert (report["leftCount"], report["rightCount"]) == (len(left), len(right))
    # Independent decimal calculation checks the requested populations/convention.
    with localcontext(Context(prec=400)):
        quantum = Decimal(1).scaleb(-places)
        mode = {"half_up": ROUND_HALF_UP, "half_even": ROUND_HALF_EVEN}[rounding]
        source_delta = sum(map(Decimal, left)) - sum(map(Decimal, right))
        output_delta = (sum(Decimal(v).quantize(quantum, rounding=mode) for v in left)
                        - sum(Decimal(v).quantize(quantum, rounding=mode) for v in right))
    assert Decimal(report["source"]["difference"]) == source_delta
    assert Decimal(report["output"]["difference"]) == output_delta
    assert report["source"]["matches"] == (abs(source_delta) <= Decimal(source_tolerance))
    assert report["output"]["matches"] == (output_delta == 0)
    assert report["validated"] == (report["source"]["matches"] and report["output"]["matches"])
    groups.append({"currency": group, "decimal_places": places,
        "source_matches": report["source"]["matches"], "output_matches": report["output"]["matches"],
        "source_difference": report["source"]["difference"], "output_difference": report["output"]["difference"]})
result = {"groups": groups}  # adapt only to the user's requested schema
output = Path("result.json")
output.write_text(json.dumps(result, ensure_ascii=False, allow_nan=False), encoding="utf-8")
assert json.loads(output.read_text(encoding="utf-8")) == result
print(json.dumps(result))
```

Choose actual output precision (0–12) and `half_up` or `half_even` from the task.
`check_totals(left, right, decimal_places, rounding="half_up", source_tolerance="0")`
returns `validated` (bool), `decimalPlaces` (int), `rounding`, `sourceTolerance`
(supplied strings), `leftCount`, `rightCount` (input lengths), `source`, `output`.
`source` and `output` each have `matches`, `difference`, `leftTotal`, `rightTotal`.
Totals/differences are strings; matches are bools; differences are left minus right.
Finite decimal strings accept signs/exponents; half-up ties round away from zero,
half-even ties to an even digit. Source tolerance is inclusive; output tolerance is zero.
Output compares sums of individually rounded values. `validated=False` is a valid
rounding-break result: report both differences; a balance claim requires True.
Source tolerance cannot excuse an output difference. These arithmetic checks do
not evaluate formulas. Use Python for full populations; keep rows out of CLI
arguments/model context.
