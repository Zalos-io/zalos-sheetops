# Compact CLI table profiles

For CLI-only analysis, `profile-table` composes the existing reader and population
validator. It scans the entire source and returns statistics without source rows.
The Python workflow in SKILL.md remains the default for scripted analysis.

```bash
sheetops profile-table --json '{"path":"source.xlsx","sheet":"Data","headerStrategy":"row","headerRow":1,"keyColumns":["reference"],"amountColumns":["amount"],"valueMode":"typed"}'
sheetops profile-table --json '{"path":"source.csv","encoding":"auto","delimiter":";","keyColumns":["reference"],"amountColumns":["amount"]}'
```

Use **normalized header names** in keyColumns/amountColumns. On missing_column,
`details.availableColumns` lists the actual names; correct the request using that
list. Amount/key checks apply to the columns selected. `maxRows` must be absent
or null; complete counts, all row widths and selected values are validated.

Success: `ok: true, data: {source, profile}`. Source contains headers, rowCount,
returnedRowCount, truncated=false and available sheet/header/CSV metadata. Profile
contains validated, rowCount, checks, exact totals[column] and columns[column]
with first/last (source order), min/max (lexical), blankCount. No rows/sourceRows
are emitted. Typed dates resolve the workbook epoch/styles. Formula evaluation,
uniqueness, matching and rounding require separate controls.

Failures retain the reader/validator's errorCode and details and exit 1. Report
the blocker; collect invalidRecordCount for each complete source before matching.
Use a Python subprocess, parse the receipt and save the requested JSON report in
the same execution. Full command input contracts are in `sheetops profile-table
--help`; other command help also exposes its own contract, without reading the
entire `schema` response. The global schema envelope is data.commands.
