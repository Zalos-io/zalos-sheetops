# Zalos SheetOps

Agent-native Excel and CSV operations toolkit.

The project prioritizes precise, deterministic spreadsheet operations for AI
agents. Features are added test-first and exposed through machine-readable CLI
contracts before agents are encouraged to rely on them.

## Agent Examples

Every command prints one line of compact JSON by default. Pass `--pretty` when
a human needs to read the output; `--raw-output` remains accepted as a
deprecated alias for the default.

```bash
sheetops doctor libreoffice
sheetops inspect --json '{"path":"output.xlsx","engine":"auto","maxFormulaCells":20}'
sheetops read-table --json '{"path":"output.xlsx","sheet":"Data","headerStrategy":"row","headerRow":1,"maxRows":50}'
sheetops cells read --json '{"path":"output.xlsx","sheet":"Summary","range":"A1:D20","include":["values","formulas","cachedValues","numberFormats"]}'
sheetops cells search --json '{"path":"output.xlsx","query":"#4BACC6","fields":["fillColors"],"sheet":"Summary","range":"A1:Z200","limit":100,"offset":0}'
sheetops hyperlinks assert --json '{"path":"output.xlsx","sheet":"Summary","checks":[{"cell":"B2","target":"https://example.com/customer","display":"Customer portal"}]}'
sheetops styles assert --json '{"path":"output.xlsx","sheet":"Merchant","checks":[{"cell":"E2","fill":"#D9EAF7"},{"cell":"H2","commentContains":"Method of Payment Code"},{"cell":"G3","fill":"#D9EAF7","formula":true}]}'
sheetops formulas scan --json '{"path":"output.xlsx","sheet":"Summary","range":"A1:D50","include":["errors"],"failFast":true,"engine":"auto"}'
sheetops normalize preview --json '{"path":"input.xlsx","sheet":"Data","headerStrategy":"detect","columns":[{"source":"Amount","kind":"amount","name":"amount"}]}'
sheetops recalc --json '{"path":"output.xlsx","outputDir":"outputs/recalc"}'
sheetops diff --json '{"left":{"path":"actual.csv"},"right":{"path":"expected.csv"},"keyColumns":["reference"],"compareColumns":["amount"],"amountTolerance":"0.00"}'
sheetops reconcile --json '{"left":{"path":"bank.csv"},"right":{"path":"processor.csv"},"leftKeyColumns":["reference"],"rightKeyColumns":["reference"],"leftAmountColumn":"amount","rightAmountColumn":"amount","amountTolerance":"0.01"}'
```

Run `scripts/setup-libreoffice` when `sheetops doctor libreoffice`
reports that LibreOffice is missing or broken. The script prints the appropriate
macOS or Linux install command and never runs `sudo` automatically.

`inspect` defaults to the XML engine so large workbook metadata checks do not
materialize every worksheet through OpenPyXL. Keep OpenPyXL read-only usage to
bounded cell or range commands where the requested area is explicit.
Inspection reports exact sheet visibility, defined names, bounded merged ranges,
and chart/image anchors. `cells search` stops at `maxCells` and reports
`stoppedEarly` rather than silently claiming a complete result.

Keep responses token-bounded: `maxRows` on table reads, `maxFormulaCells` on
inspect (default 20), `maxResults` on scans, and explicit `include` lists on
cell reads. `read-table` returns `headers` plus positional `rows` arrays and
stops scanning at `maxRows` under `headerStrategy` `row`, reporting
`"rowCount": null` with `"truncated": true` when the total is unknown.
`cells read` omits blank cells and null formulas; compare `returnedCellCount`
with `cellCount` to see how many cells were blank.

## Verification

Run these before claiming a change is ready:

```bash
rtk uv run pytest
rtk uv run ruff check .
rtk uv run ruff format --check .
```

## Releases

`v0.4.2` is the first public release, with MIT licensing and public package metadata.
It preserves the `v0.4.1` behavior baseline. It fixes worksheet bounds inspection
when dimension metadata is missing or invalid, including implicit row and cell
coordinates. This prevents non-empty worksheets from being reported as `A1`.

The `v0.4.0` release added deterministic, bounded
`cells search` across values, formulas, cached values, number formats, font
colors, and fill colors, with regex, pagination, hidden-sheet control, and
explicit scan limits. XML inspection now reports exact sheet states, defined
names, merged-range metadata, and chart/image anchors while failing closed
with `invalid_workbook` for malformed drawing relationships. Metadata output
is independently truncated without hiding total counts.

It builds on the `v0.3.1` performance and agent-output release, including the
packaging guard that keeps `zalos_sheetops.__version__` aligned with the
distribution version. The CLI dropped typer/click/rich for a zero-dependency
dispatcher, so cold start halved again and `--help` renders plain text.
`read-table` and `normalize preview` stop streaming at `maxRows` under
explicit header rows (about 80x faster on 100k-row workbooks), `validate`
runs about twice as fast, and `formulas scan` streams through the shared
expat engine instead of ElementTree. Output is compact JSON by default with
`--pretty` opt-in, `read-table` rows are positional arrays under `headers`,
`cells read` omits blank cells and null formulas and reports
`returnedCellCount`, `styles read` drops redundant row/column fields, and
`inspect` caps `formulaCells` at 20 by default (`formulaCellsTruncated`
signals more). The XML engine now expands shared formulas natively using
openpyxl's translator, so `cells read`, `formulas scan`, and
`compare labeled-cells` report the real per-cell formula instead of "=" for
shared-formula followers, and formula scans on shared-formula workbooks no
longer fall back to the slow openpyxl load. Two precision fixes validated
against 100MB+ production workbooks: `inspect` tag counting no longer drops
tags that straddle its internal 1MB scan-chunk boundaries, and `maxResults`
on `formulas scan` caps returned items across the whole workbook on both
engines instead of per sheet. The envelope contract
(`ok`/`command`/`data`/`warnings`/`artifacts`) is unchanged. Consumers upgrading from `v0.2.11` must adapt to
the new `data` payload shapes for `read-table`, `cells read`, and
`styles read`, and to the compact default output.

It includes the `v0.1.0` `inspect`, `read-table`, `validate`, and
`formulas scan` baseline, plus fixture-backed `cells read`, scoped formula
scanning, validation rules, LibreOffice doctor/recalc support, normalization,
diff, reconciliation, style/comment assertion commands, display overflow
scanning, italic style assertions, style extraction for bounded cell ranges,
and a fast XML-backed style assertion path for targeted formula, fill, and
comment checks. Formula scanning defaults to the faster XML engine, supports
formula/error category selection, supports fail-fast formula error checks,
and falls back to `openpyxl` for unsupported workbook features. Hyperlink
assertion supports cell-targeted XML validation for external or internal
targets, display text, and tooltips. Workbook inspection defaults to the XML
engine for full-workbook metadata scans, and CLI help explains command
intent, JSON payloads, and bounded range guidance.

## Install

```bash
python -m pip install zalos-sheetops==0.4.2
```

The public repository is https://github.com/Zalos-io/zalos-sheetops.
Until the initial PyPI release is available, install from its reviewed release tag.
The internal repository and its history are not part of this public distribution.

## License

MIT. Borrowed test fixtures retain their upstream licenses and attribution in
`tests/fixtures/manifest.json` and the adjacent upstream license files.
