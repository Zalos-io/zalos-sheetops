---
name: sheetops
description: Use when inspecting, validating, extracting, scanning, benchmarking, or changing SheetOps operations for Excel/CSV financial-agent work.
---

# SheetOps

Use `sheetops` for supported inspection, extraction and verification operations.
Use Polars for bulk business transformations, XlsxWriter for new reports, and
a compatible editing library for existing workbooks; SheetOps does not replace
those libraries.

Every command prints one line of compact JSON by default; pass `--pretty` only
when a human needs to read it (`--raw-output` is a deprecated alias for the
default). Do not parse human prose when structured output is available.

Keep responses token-bounded: pass `maxRows` on table reads, `maxFormulaCells`
on inspect, `maxResults` on scans, and request only the `"include"` fields you
need on `cells read` and `formulas scan`. Blank cells are omitted from
`cells read` output; `cellCount` versus `returnedCellCount` states how many
cells were blank. `read-table` returns `headers` plus positional `rows`
arrays; when it stops early at `maxRows` it reports `"rowCount": null` and
`"truncated": true`. Prefer `headerStrategy` `row` with a known `headerRow` —
it lets large scans stop early; `detect` must stream the whole sheet.

Start with:

```bash
sheetops schema
sheetops inspect --json '{"path":"workbook.xlsx"}'
sheetops read-table --json '{"path":"workbook.xlsx","sheet":"Data","headerStrategy":"row","headerRow":1,"maxRows":50}'
sheetops cells read --json '{"path":"workbook.xlsx","sheet":"Summary","range":"A1:D20","include":["values","formulas"],"engine":"xml","valueMode":"typed"}'
sheetops cells search --json '{"path":"workbook.xlsx","query":"#4BACC6","fields":["fillColors"],"sheet":"Summary","range":"A1:Z200","limit":100,"offset":0}'
sheetops styles assert --json '{"path":"workbook.xlsx","sheet":"Merchant","checks":[{"cell":"E2","fill":"#D9EAF7"},{"cell":"H2","commentContains":"Method of Payment Code"},{"cell":"G3","fill":"#D9EAF7","formula":true}]}'
sheetops validate --json '{"path":"workbook.xlsx","requiredSheets":["Data"],"requiredColumns":{"Data":["Amount"]},"uniqueKeys":{}}'
sheetops formulas scan --json '{"path":"output.xlsx","sheet":"Summary","range":"A1:D50","include":["errors"],"failFast":true,"engine":"auto"}'
sheetops verify-edits --json '{"originalPath":"original.xlsx","path":"output.xlsx","allowedRanges":{"Data":["B2:B20"]},"evaluatedPath":"output.evaluated.xlsx"}'
```

Other available commands:

```bash
sheetops normalize preview --json '{"path":"input.xlsx","sheet":"Data","headerStrategy":"detect","columns":[{"source":"Amount","kind":"amount","name":"amount"}]}'
sheetops doctor libreoffice
sheetops recalc --json '{"path":"output.xlsx","outputDir":"outputs/recalc"}'
sheetops diff --json '{"left":{"path":"actual.csv"},"right":{"path":"expected.csv"},"keyColumns":["reference"],"compareColumns":["amount"],"amountTolerance":"0.00"}'
sheetops reconcile --json '{"left":{"path":"bank.csv"},"right":{"path":"processor.csv"},"leftKeyColumns":["reference"],"rightKeyColumns":["reference"],"leftAmountColumn":"amount","rightAmountColumn":"amount","amountTolerance":"0.01"}'
```

Run `sheetops doctor libreoffice` before `sheetops recalc`. If the doctor
reports that LibreOffice is missing or broken, run `scripts/setup-libreoffice`
and follow the printed platform-specific install command.

## Totals at output precision

Before claiming that totals reconcile, compare the complete populations in the
same units, separately for each group or currency. Use `check-totals` for invoice
lines versus totals, allocations, reconciliations, report subtotals, or any other
additive equality. It checks source sums and **sums of individually rounded
values**, not just rounded grand totals.

```bash
sheetops check-totals --json '{"left":["0.104","0.104"],"right":["0.208"],"decimalPlaces":2,"rounding":"half_up"}'
```

This returns `ok: false`, `errorCode: "totals_mismatch"`, and exit code 1:
`details.source.matches` is true but `details.output.difference` is `"-0.01"`.
Successful checks put the same report in `data`, with `validated: true`.
Both comparisons must pass. `sourceTolerance` defaults to `"0"`; a justified
source tolerance never permits an output difference. Missing or invalid inputs
are errors, not zero or a pass. Empty lists are rejected; explicit `"0"` is valid.

Pass decimal **strings**, not JSON floats or already rounded display strings.
Each side accepts 1–100,000 values, up to 128 digits and exponent magnitude 128.
Choose `decimalPlaces` (required, 0–12) and `rounding` (`half_up` by default,
`half_even` when the output uses it) from the requested output convention.
Convert scaled displays such as percentages to the same units before checking.
Do not mix units, currencies, unrelated groups, or subtotal rows with their detail.
Do not use a truncated preview as a complete population.

Report source and output differences separately. Never silently change amounts,
add balancing rows, or loosen tolerance to make a check pass. Preserve source
precision and explain any unresolved rounding difference. The check is read-only
arithmetic over supplied values: it does not prove population completeness,
correct grouping, or fresh formula evaluation.

For Python workflows, use the same calculation directly:

```python
from zalos_sheetops.validation.totals import check_totals

report = check_totals(["0.104", "0.104"], ["0.208"], decimal_places=2)
assert report["validated"], report
```

## Operation Changes And Performance

Benchmark before optimizing. When adding or changing a SheetOps operation that
can scan workbook rows, cells, styles, formulas, tables, validation rules, diffs,
or reconciliation records, measure the current behavior before changing the hot
path.

Use the large-XLSX harness for inspection and validation work:

```bash
rtk uv run python -m zalos_sheetops.benchmarks.large_xlsx \
  --rows 100000 \
  --columns 12 \
  --scenarios inspect,validate,cells-search \
  --output .benchmarks/large-xlsx/results.jsonl \
  --cpu-profile-dir .benchmarks/large-xlsx/profiles
```

For a new operation, self-improve the harness before optimizing if no scenario
covers that operation yet. Update
`src/zalos_sheetops/benchmarks/large_xlsx.py` with a focused scenario and protect
the JSONL/profile contract in `tests/test_large_xlsx_benchmark.py`.

Do not claim a performance improvement without before/after evidence. Compare
elapsed time, peak traced bytes, RSS, and command summaries. Keep generated
workbooks, JSONL, and `.prof` files under `.benchmarks/` or `/tmp`.
Do not commit generated benchmark artifacts.

## CSV and Python workflows

For a complete, values-only XLSX table, use the optional `zalos-sheetops[bulk]`
Python backend after confirming the sheet and header row with a bounded XML
preview. Keep small ranges, formulas, and exact stored values on XML.

```python
from pathlib import Path
import polars as pl
from zalos_sheetops.io.read_bulk import read_bulk_table

batch = read_bulk_table(Path("source.xlsx"), "Data", header_row=1)
frame = pl.from_arrow(batch)
```

The fastexcel backend returns an Arrow RecordBatch with normalized headers and
typed values, not `read-table`'s strings. It retains interior blank rows and
text IDs, and requires explicit sheet/header coordinates. It rejects formulas,
Excel errors, mixed column types and unsafe large numeric values. For a justified
mixed text column, pass `column_types={"normalized header": "string"}`; this adds
an XML precision check and may be slower. Never coerce away an error to get a
result. Use XML to investigate rejected inputs. Formula evaluation and output
verification remain separate requirements. Bulk reads materialize the sheet;
use XML for bounded reads or when memory is constrained.
Text IDs already remain strings without overrides. Read the table once and
reuse the frame for inspection, transformation and control totals.

```bash
sheetops read-table --json '{"path":"source.csv","encoding":"auto","delimiter":"\t","maxRows":20}'
```

CSV reads preserve string values and reject extra fields. `auto` detects BOMs
and otherwise uses UTF-8; legacy encodings require an explicit name. The default
delimiter is comma, or tab for `.tsv`. `sourceRows` gives physical record start
lines; `truncated` marks a partial preview, `returnedRowCount` counts returned
records, and `rowCount` is null until the full population is read. Headers are
normalized. CSV sources in `diff` and `reconcile` accept
`encoding`, `delimiter`, and `headerRow` too. Do not treat a preview as a full
reconciliation. Preserve long IDs and leading zeros, parse amounts/dates explicitly,
and reject ambiguous duplicate keys instead of silently choosing a row.

Create and edit in separate paths: stream row-ordered new detail sheets with
XlsxWriter, use buffered writes for layouts that revisit rows, and never rebuild
a formatted existing workbook through a DataFrame writer. Verify edits against
the original with an explicit allowed scope. Evaluate formula outputs in an
actual calculation engine and check independent expected results; cached values
and successful serialization do not prove calculation or business correctness.
