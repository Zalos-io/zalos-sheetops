---
name: sheetops
description: Use when reading, validating or changing Excel/CSV data with SheetOps.
---

# SheetOps

## Data-analysis workflow

Use the documented Python APIs below directly on the supplied files. This task
uses an installed spreadsheet library; a successful import in the analysis
script checks availability. The validated read supplies the data evidence.

1. Start the analysis script with the documented imports and the actual reader.
   Use given paths, sheets, header rows, encodings and delimiters. Inspect once
   only when a required source coordinate is unknown.
2. Validate each entire required population, then compute the requested controls
   on the loaded rows. Reuse the receipts and rows for all report fields.
3. Construct and save the requested JSON in that same execution, including its
   requested blocked shape when a control rejects. Print compact findings.

Finish when the requested report is saved and every applicable control has
passed or its blocker is reported. A successful import plus these receipts
completes availability/data checks. Extra package browsing or manual parser
reimplementation adds no control. Use a heredoc or a writable working directory.

For an unexpected API/import error, resolve the specific gap with the callable's
signature/docstring or that command's `--help` in one lookup, then use the actual
reader/validator. Inspect implementation for a reproducible contradiction or
package bug. Invalid input is a reportable blocker, not an API-discovery task.
All tools remain available for a justified investigation; preserve rejection.
Python imports `zalos_sheetops`; CLI: `sheetops`.

## Complete tables: read → profile → report

```python
import json
from pathlib import Path
from zalos_sheetops.io.read_table import read_table
from zalos_sheetops.io.read_csv_table import read_csv_table
from zalos_sheetops.io.read_csv_profile import read_csv_profile
from zalos_sheetops.validation.table_profile import profile_table
from zalos_sheetops.errors import SheetOpsError

try:
    # Choose one reader; None reads the entire population.
    table, receipt = read_csv_profile(Path("source.csv"), key_columns=["reference"],
        amount_columns=["amount"], encoding="auto", delimiter=",")
    # XLSX alternative: read_table(..., max_rows=None), then profile_table(...).
    # Adapt fields/shape to the requested report.
    result = {"row_count": receipt["rowCount"], "totals": receipt["totals"]}
except SheetOpsError as error:
    # Adapt this blocker to the user's requested JSON shape in this same run.
    result = {"status": "blocked", "errorCode": error.error_code, "details": error.details}
Path("result.json").write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
print(json.dumps(result, ensure_ascii=False))
```

Readers return `headers` (stripped, lowercased, whitespace collapsed), `rows`, `rowCount`, `returnedRowCount`,
`truncated`. XLSX rows are string lists; CSV rows are dictionaries (CSV also adds
`sourceRows`, `encoding`, `delimiter`). Preserve IDs, amounts, whitespace and
multiline values. Use returned header names (e.g. `reference`, not `Reference`)
in key/amount lists. Resolve a missing column from the loaded `table["headers"]`.
Legacy encodings need an explicit codec; specify CSV delimiters.

`profile_table` checks the complete population, counts, rectangular rows,
nonblank text keys and finite bounded decimal-string amounts. It returns
`validated`, `rowCount`, exact `totals[column]`, and `columns[column]` containing
`first`/`last` (source order), `min`/`max` (lexical), `blankCount`.
Use the requested statistic and loaded rows for other fields. Previews
(`maxRows` CLI / `max_rows` Python) cannot validate totals.

**Report a rejection from its receipt.** `table_width_mismatch` detects populated
unheaded cells, including zero; `details["sheet"]` and `details["row"]` locate the
first offending source row. `invalid_table_values` gives `invalidRecordCount`,
`issueCount`, up to 20 `issues` with record/source-row locations, `issuesTruncated`.
Each issue includes `valueType`; strings add the original `valuePreview` (up to
80 characters) and `valueTruncated`. These are bounded examples; invalid-record
counts still cover the entire population. Use this evidence and the loaded rows
to explain the rejection; investigate another source read only for a concrete gap
or conflicting evidence.
Profile every whole source and collect all invalid-record counts before matching
or filtering. Investigate conflicting evidence or additional required controls;
a documented rejection alone needs no source-code audit. Never discard bad rows.
CLI: `ok: true, data` or `ok: false, errorCode, message, details` with nonzero exit.

For CSV source validation, `read_csv_profile(path, *, key_columns, amount_columns,
encoding="auto", delimiter=",", header_row=1)` returns `(complete_table, receipt)`.
It reuses the same complete-population/value checks, rejects row limits, and adds
`receipt.source`: actual path, SHA-256, headers, encoding, delimiter and row count.
Rejections preserve their codes/counts and include `details.sourceExcerpts` read
independently from original CSV text (up to 3 physical lines / 240 characters per
retained source row, with truncation flags). A changed source blocks the receipt.
Use this source evidence in the read/validate/report program; resolve additional
reads only for a concrete gap or conflicting evidence.

## Dates and bounded cell evidence

For dates, add `value_mode="typed"` to `read_table`: workbook epoch (1900/1904)
and date styles are applied; IDs/amounts retain source strings. Date cells become
datetime strings. With `from datetime import datetime`, a requested date-only
field is `datetime.fromisoformat(value).date().isoformat()`; otherwise keep time.
Raw mode returns serials; typed mode already resolves the epoch.
Typed XML table/cell receipts include `valueMode: "typed"` and `dateSystem:
"1900" | "1904"`. These identify the applied conversion; combine them with the
returned values/formats for date evidence. Investigate only a concrete conflict,
rather than opening workbook XML to redo a successful conversion.


For known coordinates, use this Python API in the same kernel as validation and
reporting. It returns the data dictionary directly (no CLI `ok`/`data` envelope):

```python
import hashlib
import json
from datetime import datetime, timedelta
from decimal import Decimal
from pathlib import Path
from openpyxl.styles.numbers import is_date_format
from openpyxl.utils.cell import get_column_letter, range_boundaries
from zalos_sheetops.io.read_cells import read_cells

# Bind actual sheet/range, column order, header labels and requested record count.
source, sheet, bounds = Path("source.xlsx"), "Data", "A1:C7"
expected_headers = ["reference", "amount", "date"]
expected_records = 3  # supplied business expectation; otherwise use actual population
before = hashlib.sha256(source.read_bytes()).hexdigest()
data = read_cells(source, sheet, bounds, engine="xml",
                  value_mode="typed", include=["values", "sourceValues", "formulas",
                                                "cachedValues", "numberFormats", "dataTypes"])
min_col, top, max_col, bottom = range_boundaries(bounds)
columns = [get_column_letter(col) for col in range(min_col, max_col + 1)]
cells = {cell["address"]: cell for cell in data["cells"]}
assert len(cells) == data["returnedCellCount"] == len(data["cells"])
assert data["rowCount"] == bottom - top + 1 and data["columnCount"] == len(columns)
assert data["cellCount"] == data["rowCount"] * data["columnCount"]  # includes all blanks
assert set(cells) <= {f"{col}{row}" for row in range(top, bottom + 1) for col in columns}
assert data["dateSystem"] in {"1900", "1904"}
assert [cells[f"{col}{top}"]["value"] for col in columns] == expected_headers
records, empty_rows = [], []
for row in range(top + 1, bottom + 1):
    addresses = [f"{col}{row}" for col in columns]
    if all(address not in cells for address in addresses):
        empty_rows.append(row)
        continue
    assert all(address in cells and "formula" not in cells[address] for address in addresses)
    reference_cell, amount_cell, date_cell = [cells[address] for address in addresses]  # adapt order
    reference, amount = reference_cell["sourceValue"], amount_cell["sourceValue"]
    assert isinstance(reference, str) and reference.strip()
    assert isinstance(amount, str) and Decimal(amount).is_finite()
    parsed = datetime.fromisoformat(date_cell["value"])
    assert date_cell["dataType"] == "n" and is_date_format(date_cell["numberFormat"])
    # Independent calendar-day control from the original serial, no float conversion.
    serial = Decimal(date_cell["sourceValue"])
    assert serial.is_finite() and serial >= 1
    days = int(serial)
    epoch = datetime(1904, 1, 1) if data["dateSystem"] == "1904" else datetime(1899, 12, 30)
    if data["dateSystem"] == "1900" and days < 60:
        days += 1  # Excel's fictitious 1900 leap day
    assert parsed.date() == (epoch + timedelta(days=days)).date()
    records.append({"reference": reference, "amount": amount, "date": parsed.date().isoformat()})
assert len(records) == expected_records
assert 1 + len(records) + len(empty_rows) == data["rowCount"]
assert hashlib.sha256(source.read_bytes()).hexdigest() == before
result = {"records": records}  # adapt requested fields; retain datetime if requested
output = Path("result.json")
output.write_text(json.dumps(result, ensure_ascii=False, allow_nan=False), encoding="utf-8")
assert json.loads(output.read_text(encoding="utf-8")) == result
print(json.dumps({"dateSystem": data["dateSystem"], "result": result}))
```

Use the include names above. `sourceValue` retains original numeric XML strings
(including every digit) or decoded text; `value` supplies typed ISO dates.
The receipt's `dateSystem` and number formats establish date interpretation;
use this evidence directly. Formula `sourceValue` is a saved cache, not a fresh
calculation. This evidence comes from the same reader; independent business
controls still compare actual source populations or expected business results.
`rowCount`/`columnCount` span the full requested grid, including headers and blank
rows. `cellCount = rowCount * columnCount`; `returnedCellCount = len(cells)` counts
populated cells (headers, zero, False and formulas included). Neither is a record
count: a 7×3 range is 21 cells; its header plus three complete data rows returns 12.
Iterate all physical rows, skip only wholly blank rows and reject partial rows.
Account for headers + records + empty rows; no regex or date-parser reconstruction
is needed. The example checks calendar dates; preserve and check time separately
when requested. For ISO date cells (`dataType="d"`), compare source/typed ISO dates
instead of numeric serial arithmetic. Reads are bounded to 10,000 cells. Typed mode uses
styles/workbook epoch and emits ISO strings; convert ISO datetime to the requested
date shape. Raw mode retains serials; table reads default to raw, so explicitly
select `value_mode="typed"` for dates. A bare number alone is not date evidence.
Use `sourceValue` for exact numeric amounts rather than converting `value` through
float. The XML/auto engines support `sourceValues`; openpyxl rejects that option.

Formula text and saved cached results are separate; caches may be stale or absent.
Resolve required missing cells/caches. Recalculation requires an engine, not a
cache read. Include source sheet/row only where the requested schema permits.
Keep exact output schemas unchanged. For unknown coordinates use one bounded
preview; for read-only reports the recipes here supply the workflow.


CLI equivalent: `sheetops cells read --json` with path, sheet, range, engine,
valueMode and `"include"` fields. For unknown structure use a bounded inspection:
`sheetops inspect --json '{"path":"source.xlsx","maxFormulaCells":10}'`.

## Financial CSV controls

For CSV transaction matching, use the public complete workflow in the analysis
program; it profiles both sources before matching and independently verifies
classification, exact differences, unique assignments and both-side coverage:

```python
from pathlib import Path
from zalos_sheetops.finance import reconcile_csv
analysis = reconcile_csv(Path("left.csv"), Path("right.csv"),
    left_key_columns=["reference"], right_key_columns=["posting reference"],
    left_amount_column="amount", right_amount_column="net amount",
    amount_tolerance="0", left_encoding="auto", right_encoding="auto",
    left_delimiter=",", right_delimiter=",")
if analysis["status"] == "blocked":
    result = {"status": "blocked", "invalid_rows": analysis["invalidRecordCount"]}
else:
    result = {"summary": analysis["report"]["summary"]}  # adapt requested fields
```

For grouped source/per-line-rounded totals, import `check_grouped_totals_csv`
from the same module. Its keyword arguments are `group_column`, `side_column`,
`amount_column`, `decimal_places` (the full group→integer precision mapping),
`left_label`, `right_label`, `rounding="half_up"`, `source_tolerance="0"`,
`encoding="auto"`, `delimiter=","`. It checks every record and group assignment,
then independently verifies source and sums-of-rounded-lines totals. Unknown
side/group values fail; no record is filtered away. `half_even` is also supported.
`report.groups` is sorted by `group`; entries contain the `check_totals` report
plus the source group string. Preserve both differences and every false match flag.

Both workflows return `status` (`checked` or `blocked`), `report`, `sources`,
`rejections`, `invalidRecordCount` and `checks`. Sources are named `left`/`right`
for reconciliation and `source` for totals. Each has `path`, `rowCount`,
`invalidRecordCount`, `sha256Before`, `sha256After`, `headers`, `encoding`,
`delimiter`, `headerRow`. Rejections retain complete counts and bounded original
source excerpts; blocked `report` is None. A checked reconciliation report has
`summary` plus `matched`, `variance`, `leftUnmatched`, `rightUnmatched` lists.
Right-unmatched keys use right column names; other keys use left names.

Keep full reports in Python; map only the requested result schema, save/parse it
in the same execution and bind its hash and source counts/hashes in the requested
evidence receipt. Add independent external business expectations. Mechanical
checked status is not financial balance, formula evaluation or posting approval.
Expected input rejection needs no source/API discovery or unchecked fallback.
For an unavailable workflow import, use the equivalent lower-level recipe in
[finance-python-legacy.md](references/finance-python-legacy.md) or report setup
unavailable; preserve all checks. XLSX inputs retain complete table reads and
must use equivalent lower-level controls.

- For Arrow/fastexcel values-only bulk reads, read [bulk.md](references/bulk.md).
- For editing, styles, normalization, diff or formula delivery, read
  [workbook-operations.md](references/workbook-operations.md). Preserve unrelated
  cells; formula delivery requires actual successful engine evaluation.
- Before changing or benchmarking operations, read [development.md](references/development.md).

- For CLI-only complete-table profiling, read [cli-profile.md](references/cli-profile.md).
