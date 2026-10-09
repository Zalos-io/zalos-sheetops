# SheetOps Architecture And Reference Repositories

This document is the orientation file for future AI agents contributing to
`zalos-sheetops`. It explains the package architecture, the trust model, and
the external repositories we studied or borrowed fixtures from.

## Product Goal

`zalos-sheetops` is an agent-native Excel and CSV toolkit for Codex, Claude
Code, and financial-agent workflows. Its job is to make common spreadsheet
operations precise, deterministic, and safe enough that agents can call a
tested command instead of rewriting spreadsheet logic inside every financial
agent.

The package favors a small set of trusted operations over a large set of
loosely-tested helpers.

## Architecture

The public surface is a JSON-first CLI plus small Python library modules behind
each command.

```text
src/zalos_sheetops/
  cli.py                 # JSON CLI dispatcher
  command_registry.py    # machine-readable command manifest
  contracts.py           # success/error response helpers
  errors.py              # stable SheetOpsError
  paths.py               # path hardening and local file validation

  io/
    inspect.py           # workbook/sheet/cell metadata
    search_cells.py      # bounded, paginated value/formula/style search
    headers.py           # header detection and canonicalization
    read_table.py        # XLSX table extraction
    read_csv_table.py    # planned CSV table extraction

  validation/
    workbook.py          # required sheets/columns and uniqueness checks

  normalize/
    text.py              # header/text canonicalization
    amounts.py           # planned amount normalization
    dates.py             # planned date normalization
    preview.py           # planned normalize-preview command logic

  formulas/
    scan.py              # formula and formula-error discovery

  recalc/
    libreoffice.py       # planned LibreOffice recalculation wrapper

  diff/
    tabular.py           # planned table diff helpers

  reconcile/
    exact.py             # planned exact reconciliation helpers
```

## Command Contract

Every command returns structured JSON. Agents should parse these fields and
avoid parsing human prose.

Success:

```json
{
  "ok": true,
  "command": "inspect",
  "data": {},
  "warnings": [],
  "artifacts": []
}
```

Failure:

```json
{
  "ok": false,
  "command": "inspect",
  "errorCode": "file_not_found",
  "message": "Workbook file does not exist",
  "details": {},
  "warnings": [],
  "artifacts": []
}
```

## Trust Model

- Public behavior starts with a failing test.
- Tests must verify exact JSON shapes for CLI behavior.
- Real workbook behavior should use fixture-backed tests where possible.
- Fixture provenance is mandatory for borrowed files.
- Commands fail closed when inputs are ambiguous.
- Source files are not mutated by default.
- Recalculation and generated outputs must go to explicit output paths.

## Fixture Strategy

Fixtures live under `tests/fixtures/`.

```text
tests/fixtures/
  README.md
  manifest.json
  upstream/
    calamine/
    fastexcel/
    openpyxl/
    xlsx2csv/
    xlsxwriter/
```

Every upstream fixture must be registered in
`tests/fixtures/manifest.json` with:

- `fixtureId`
- `sourceRepo`
- `sourcePath`
- `license`
- `destination`
- `sha256`
- `behavior`

The manifest hash-locks fixture files so agents cannot unknowingly change test
ground truth.

## Reference Repositories

These repositories were used as design or fixture references.

### Spreadsheet Agent Skill

| Repo | Used For |
| --- | --- |
| `anthropics/skills` | `xlsx` skill conventions: formula-first workbooks, LibreOffice recalc, formula-error scans, color conventions, and agent-facing spreadsheet workflows |
| `dust-tt/dust@dc59bd33` (`front/lib/api/sandbox/image/profile/soffice/xlsx_inspect.py`) | MIT-licensed design reference for drawing-aware inspection, defined names, displayed color resolution, and paginated workbook search. No upstream implementation or tests were copied; SheetOps behavior was specified and tested independently. |

### Core Workbook IO And Writing

| Repo | License | Used For |
| --- | --- | --- |
| `openpyxl` | MIT | Existing workbook inspection, formula text vs cached values, hidden sheets, hyperlinks, comments, defined names, formula token/translation behavior, XLSX fixtures |
| `XlsxWriter` | BSD-2-Clause | Generated workbook guardrails, formula writing, tables, data validation, conditional formatting, overlap checks, comparison XLSX fixtures |
| `fastexcel` | MIT | Fast Excel ingestion patterns, `header_row`, `skip_rows`, whitespace, duplicate columns, sheet visibility, table-offset fixtures |
| `calamine` | MIT | Cross-format spreadsheet parsing behavior, formula edge cases, shared formulas, `.xls`/`.xlsb`/`.ods` future support, fixtures |

### Data Processing

| Repo | Used For |
| --- | --- |
| `polars` | DataFrame transform patterns: `select`, `with_columns`, `filter`, `group_by`, `join`, `concat`, lazy scans |
| `pandas` | Excel/CSV IO option reference: sheet selection, `header`, `dtype`, `parse_dates`, `thousands`, `decimal`, `skipfooter` |
| `numpy` | Numeric behavior reference only; not a primary SheetOps dependency |

### Validation And Data Contracts

| Repo | Used For |
| --- | --- |
| `pandera` | Schema/check design: required columns, strict schemas, uniqueness, nullable/coerce behavior, dataframe-level checks |
| `frictionless-py` | File/table diagnostics: dialect, header rows, missing values, checklist-style validation |
| `great_expectations` | Expectation-style validation concepts and reporting vocabulary |

### Matching And Reconciliation

| Repo | Used For |
| --- | --- |
| `RapidFuzz` | Future fuzzy matching, string similarity thresholds, scorer behavior |

### CSV, Diff, And Conversion

| Repo | License | Used For |
| --- | --- | --- |
| `xlsx2csv` | MIT | XLSX-to-CSV fixtures and expected CSV baselines for datetimes, UTF-8, empty rows, hyperlinks, delimiters |
| `csvkit` | MIT | CSV dialect and cleanup edge cases: BOM, Latin-1, short/long rows, no-header files, newline behavior |
| `csv-diff` | Apache-2.0 | Keyed row diff shape: added rows, removed rows, changed rows |
| `daff` | MIT | Table diff concepts, row/column alignment, patch-style comparison behavior |

### Formula Evaluation And Recalculation

| Repo | Used For |
| --- | --- |
| `xlcalculator` | Offline formula-evaluation reference and limitations |
| `formulas` | Formula parser/evaluator reference, supported function landscape, compare-method ideas |
| LibreOffice | Primary future recalculation path because it recalculates real XLSX workbooks instead of approximating Excel behavior |

## Borrowed Fixture Sources

The current fixture manifest borrows files from:

- `fastexcel`
- `calamine`
- `openpyxl`
- `xlsx2csv`
- `XlsxWriter`

See `tests/fixtures/manifest.json` for exact source paths, destination paths,
licenses, SHA256 hashes, and the behavior each fixture protects.

## How To Add A New Reference Or Fixture

1. Prefer a small synthetic fixture if it proves the behavior.
2. Borrow an upstream fixture only when real workbook structure matters.
3. Copy the file into `tests/fixtures/upstream/<repo>/`.
4. Record source repo, source path, license, destination, SHA256, and protected
   behavior in `tests/fixtures/manifest.json`.
5. Add or update a test that fails before the implementation change.
6. Run:

```bash
rtk uv run pytest
rtk uv run ruff check .
rtk uv run ruff format --check .
```

