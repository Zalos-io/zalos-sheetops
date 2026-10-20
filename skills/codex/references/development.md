# SheetOps operation development

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
