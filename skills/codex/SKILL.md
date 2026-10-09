---
name: sheetops
description: Use when inspecting, validating, extracting, scanning, benchmarking, or changing SheetOps operations for Excel/CSV financial-agent work.
---

# SheetOps

Prefer `sheetops` before writing custom spreadsheet code.

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
sheetops cells read --json '{"path":"workbook.xlsx","sheet":"Summary","range":"A1:D20","include":["values","formulas"]}'
sheetops cells search --json '{"path":"workbook.xlsx","query":"#4BACC6","fields":["fillColors"],"sheet":"Summary","range":"A1:Z200","limit":100,"offset":0}'
sheetops styles assert --json '{"path":"workbook.xlsx","sheet":"Merchant","checks":[{"cell":"E2","fill":"#D9EAF7"},{"cell":"H2","commentContains":"Method of Payment Code"},{"cell":"G3","fill":"#D9EAF7","formula":true}]}'
sheetops validate --json '{"path":"workbook.xlsx","requiredSheets":["Data"],"requiredColumns":{"Data":["Amount"]},"uniqueKeys":{}}'
sheetops formulas scan --json '{"path":"output.xlsx","sheet":"Summary","range":"A1:D50","include":["errors"],"failFast":true,"engine":"auto"}'
```

Next milestone commands:

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

fall back to custom code only when `sheetops schema` shows no supported
command for the required operation, or when `sheetops` returns a structured
error that explicitly identifies unsupported behavior.
