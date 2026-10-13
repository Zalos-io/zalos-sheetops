# Zalos SheetOps

Agent-native Excel and CSV operations toolkit.

The project prioritizes precise, deterministic spreadsheet operations for AI
agents. Features are added test-first and exposed through machine-readable CLI
contracts before agents are encouraged to rely on them.

Typed XML reads, `verify-edits`, and enforced delivery are 0.4.4 candidate features on this branch.
The published PyPI release remains 0.4.2.

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

For calendar dates and times, opt into `"valueMode": "typed"` on XML cell
or table reads. It applies Excel's 1900/1904 date system and cell formats;
cell results are ISO strings, while tables retain string rows. Raw XML
reads remain the default, including exact stored serials. OpenPyXL reads
retain their existing typed behavior.

For edits, preserve the original, write a staged output, and declare the
allowed ranges. `verify-edits` fails on changed cells, styles, comments or
hyperlinks outside those ranges and on undeclared structure changes. It also
checks chart/drawing, embedded and theme parts conservatively. Local names,
print titles and resolved dimension/conditional styles are compared as well.
Array formula footprints must fit the declared ranges, including when removed.
New formulas require
an `evaluatedPath` containing a copy evaluated separately by Excel or
LibreOffice; missing/error caches, mismatched formulas and unsafe modern
formula storage fail verification. Array output caches must cover the entire
declared footprint. Changed data-table formulas are rejected as unsupported;
unchanged data tables are preserved. The command checks the supplied results;
it cannot establish cache freshness or prove the task's expected answer.
Independently check expected values and keep the evaluation copy separate
from the deliverable, since spreadsheet engines may rewrite formatting.

```bash
sheetops verify-edits --json '{"originalPath":"original.xlsx","path":"output.xlsx","allowedRanges":{"Data":["B2:B20"]},"evaluatedPath":"output.evaluated.xlsx"}'
```

The edit guard loads both workbooks through OpenPyXL. It supports cell edits
and explicitly declared structural changes within existing sheets; it does
not certify workbook renames, VBA execution or every unsupported Excel feature.

## Verification

Run these before claiming a change is ready:

```bash
rtk uv run pytest
rtk uv run ruff check .
rtk uv run ruff format --check .
```

## Public benchmarks

The [October 2026 baseline](benchmarks/results/2026-10-09/baseline.md) includes
480 reader measurements, 33 paired task attempts with 30 usable pairs, audited
native Excel grades, and a public archive of failed cases. Both task arms passed
20/30; this run does not establish an overall app-level speed or accuracy win.
SheetOps accelerated deep scalar reads; that baseline exposed typed-date,
preservation and native formula validation gaps.

The [October 10 rerun](benchmarks/results/2026-10-10/report.json) completes another
480 reader measurements and all 66 agent attempts under the same recorded
controls. Both SheetOps reader interfaces improve from **54/96 to 96/96** checks
passed, resolving all 42 typed-date failures per interface. The 14 separate
feature contracts and the 30-task golden self-check also pass.

| 400,000-row workload | OpenPyXL median | SheetOps CLI median |
| --- | ---: | ---: |
| Typed bounded read near the end | 11.71 s | 5.08 s |
| Typed bounded read near the start | 0.068 s | 0.133 s |
| Full table read | 13.91 s | 8.90 s |

These are three fresh-process repetitions on one host with uncontrolled OS file
cache. Deep bounded reads use 35.3 MiB peak CLI RSS versus 66.9 MiB for OpenPyXL;
full-table reads use 663.0 MiB versus 491.3 MiB. RSS is the maximum for one process,
not aggregate memory. Calamine and fastexcel remain faster for full-table reads.

The native Excel task comparison is complete: **15/30 baseline and 14/30
candidate**, versus 20/30 for each arm in the earlier run. Among the 11 pairs
where both arms passed, the median baseline/candidate time ratio is 0.849;
this subset favors baseline and has selection bias. Each run has one agent
repetition per task; stochastic inference and provider revisions are uncontrolled,
so these scores do not isolate a regression or establish an overall app win.

Excel calculated and saved 30 formula-containing output copies and rejected
three with content-repair warnings; repair was declined. All three were from
timed-out attempts, retained as timeouts with invalid workbook status. The
native audit's three not-ready copies are these rejections, with zero outstanding
calculation actions. All 88 input/output sources remain unchanged; 30 golden
calculations were reused only after hash/version checks and a passing self-check.
All attempts, controls, hashes, the [native failure archive](benchmarks/results/2026-10-10/failure-reproductions.zip)
and the [secondary LibreOffice archive](benchmarks/results/2026-10-10/libreoffice-failure-reproductions.zip)
are retained. LibreOffice secondary scores also total 15/30 and 14/30, but
individual grades differ. Both terminal images lack a calculation engine, and
agent timing excludes external grading.

The candidate used `verify-edits` in 21/33 attempts, with 23 rejecting responses
and eight passing responses. Guard alerts are distinct from invalid command
arguments. Its declared edit ranges must match the requested cells and conditions;
a passing guard does not prove expected answers, requested formula presence or
the final delivered artifact. In that instruction-only cohort, agents could omit
the guard or submit after rejection.

The [October 11 delivery-enforced rerun](benchmarks/results/2026-10-11/report.json)
retains all 66 attempts and uses the same 30 eligible task pairs. Both arms share
the independent parent delivery policy and a trusted host LibreOffice engine;
the baseline terminal has no SheetOps CLI. Agent timing now includes runtime
validation and formula calculation. Four pairs run concurrently, whereas the
earlier cohorts were sequential, so historical wall times are not comparable.

| Native Excel outcome | Baseline | SheetOps candidate |
| --- | ---: | ---: |
| Correct, approved deliveries | 6/30 | 9/30 |
| Approved deliveries, including wrong answers | 10/30 | 12/30 |
| Blocked deliveries | 20/30 | 18/30 |
| Correct raw staging files, including failed attempts | 16/30 | 16/30 |
| Median time, all attempts | 103.9 s | 129.8 s |

Only four pairs delivered correct answers in both arms. Their median
baseline/candidate time ratio is 0.514, meaning the candidate took about 1.95 times
as long in that small, selected subset. This single stochastic cohort establishes
neither a stable speed win nor an overall accuracy improvement. Seven approved
workbooks still fail the task rubric: frozen declared scope and formula validity
do not prove that the agent understood the request correctly.

Fresh Excel 16.113.4 references pass the 30/30 self-check. Excel saved 24 actual
formula output copies and rejected five with content-recovery warnings; recovery
was declined. Those five are blocked failures, not delivered artifacts. All 87
sources are unchanged, and no native calculation actions remain outstanding.
Secondary LibreOffice grades are retained separately. The
[failure archive](benchmarks/results/2026-10-11/failure-reproductions.zip) contains
public task inputs and hashed raw failure outputs; blocked workbooks are diagnostic
copies. The cohort measures frozen Desktop agent core commit `acff34f`, not
Electron UI end-to-end. Later sparse native-probe and Kanban history fixes have
separate regression tests and are identified in the report rather than attributed
to the frozen timing measurements.

The benchmark harness keeps correctness and speed separate. It uses pinned
public inputs, fresh reader processes, full-output digests, and independent
OpenPyXL reference reads. Optional reader dependencies stay in the `benchmark`
dependency group; normal package installation does not install them.

```bash
rtk proxy uv sync --group benchmark --group dev
rtk proxy uv run python -m zalos_sheetops.benchmarks.public.runner prepare --data /tmp/sheetops-bench/data --output /tmp/sheetops-bench/results
rtk proxy uv run python -m zalos_sheetops.benchmarks.public.runner select-tasks --data /tmp/sheetops-bench/data --output /tmp/sheetops-bench/results
rtk proxy uv run python -m zalos_sheetops.benchmarks.public.runner operations --data /tmp/sheetops-bench/data --output /tmp/sheetops-bench/baseline --sizes 25000,100000,400000 --repeats 3
rtk proxy uv run python -m zalos_sheetops.benchmarks.public.contracts --data /tmp/sheetops-bench/data --output /tmp/sheetops-bench/contracts
rtk proxy uv run python -m zalos_sheetops.benchmarks.public.runner replay --data /tmp/sheetops-bench/data --output /tmp/sheetops-bench/replay --case 25000-cells-start --adapter sheetops-cli
```

The range benchmark includes both typed ranges containing dates and scalar
ranges without the date column, at the beginning, middle and end of each sheet.
Formula cases also require formula text and cached values. Unsupported results
remain visible; a fast incorrect output never enters the timing comparison.
The feature checks cover selected metadata, merges, number formats, hyperlinks,
and exact decimal reconciliation, not every Excel feature.

Task selection keeps the initial seeded 30 tasks and records three reviewed
reference conflicts. It adds three tasks of the same type from the same seeded
stream, producing 33 attempted pairs and 30 usable pairs. Pass `--limit 33` to
the agent runner; its default is 30. The selection manifest explains exclusions
without removing their original attempts.

Paired Zalos tasks use the installed desktop agent through its headless CLI.
Build the two images with `agents build-images --dockerfile
scripts/benchmarks/Dockerfile --skill-root PATH_TO_BASELINE_XLSX_SKILL --output
RUN_DIRECTORY`. The image builder expects the existing skill's helper scripts.
Then run `agents run` with `--output RUN_DIRECTORY`, `--dataset` pointing to
`data/verified/spreadsheetbench_verified_400`, `--tasks results/tasks.json`,
`--launcher`, `--env-file`, `--baseline-skill`, `--sheetops-skill`, and
`--soffice`. Both commands use the module
`zalos_sheetops.benchmarks.public.agents`. The environment file provides the
AI Gateway credential only to inference, never to the Docker terminal.

The runner fixes model, provider, reasoning, task order and sandbox limits;
records image IDs, input and skill hashes; alternates arm order; and resumes
only matching configurations. Each arm gets a new workspace and home. Only
the input workbook is mounted, and golden workbooks stay outside the agent's
sandbox. A timeout is retained as a measured outcome.

Regrade saved outputs without another model run:

```bash
rtk proxy uv run python -m zalos_sheetops.benchmarks.public.regrade --dataset /tmp/sheetops-bench/data/verified/spreadsheetbench_verified_400 --tasks /tmp/sheetops-bench/results/tasks.json --runs /tmp/sheetops-bench/agents --output /tmp/sheetops-bench/agents/agents-graded.jsonl
rtk proxy uv run python -m zalos_sheetops.benchmarks.public.report --root /tmp/sheetops-bench --output /tmp/sheetops-bench/report --agent-records /tmp/sheetops-bench/agents/agents-graded.jsonl
```

Regrading verifies recorded input/output hashes and preserves raw records.
Pass `--soffice PATH_TO_LIBREOFFICE` if evaluation copies are not already
available. For native Excel validation, run the `native` module's `prepare`
action with `--dataset`, `--tasks`, `--runs`, `--output` and `--excel-version`.
Open each queued copy in Excel, recalculate, save, then close it. Decline any
content repair and record that rejection in the copy's manifest entry as
`nativeRejection`. Run `native verify --output EVALUATION_ROOT`, then regrade
with `--native-evaluation-root EVALUATION_ROOT`. The audit requires intact
sources and saved formula caches; the grader checks evaluated file hashes.
Keep LibreOffice scores separately to measure interoperability differences.

Recalculation changes only evaluation copies. Agent scores cover
answer values, explicit requested styles and preservation; they are not the
official SpreadsheetBench leaderboard score. This experiment compares the
package **plus** spreadsheet skill configuration and does not measure Electron
UI behavior or isolate the package's contribution from the skill's.
Agent floating-point values use Excel's 15 significant digits; integer IDs,
strings, reader digests and decimal reconciliation retain exact comparison.

Public sources: [excel-read-benchmark](https://github.com/nazmul284/excel-read-benchmark)
(MIT; adapted mixed-column generator), [ExcelBench](https://github.com/SynthGL/ExcelBench)
(MIT; selected feature fixtures), and
[SpreadsheetBench Verified-400](https://huggingface.co/datasets/KAKA22/SpreadsheetBench)
(CC-BY-SA-4.0; selected tasks). Exact revisions, URLs and SHA256 hashes are in
`data/sources.json`; the Verified-400 archive is checked against its pinned
SHA256 before extraction. Upstream inputs retain their own licenses. Redistributed
SpreadsheetBench workbooks, task text and adapted dataset content must retain
CC-BY-SA-4.0 attribution and licensing; they are not relicensed under this
package's MIT license.

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

## Runtime delivery policy (candidate)

The `sheetops-delivery` Hermes plugin validates staged exports outside the model
loop. Enable it with `plugins.enabled` and `plugins.required_delivery` containing
`sheetops-delivery`, and `plugins.delivery_extensions: [".xlsx"]`. The latter denies
unapproved exports even when preparation fails. Zalos Desktop applies these
settings automatically. A matching runtime with delivery hooks and guarded file
routes is required; installing the package alone cannot enforce downloads.
The runtime invokes `finish_delivery` even on cancellation to close prepared
policy state without validation or approval. The plugin also accepts the older
`post_llm_call` cleanup callback for compatible hosts.

Before editing, the agent writes the per-turn contract named in runtime context
in a separate tool call. Its `source`, new `output`, and bounded `allowedRanges`
freeze before the next file tool runs. The runtime snapshots the input, validates
the actual output against that scope, evaluates all formulas, and refreshes only
formula caches in the delivered OOXML package. Missing engines, evaluation errors,
changed contracts, and unmatched bytes block delivery. The scope declaration does
not prove that the agent interpreted the user's request correctly.

Set `plugins.entries.sheetops-delivery.settings.calculation_engine` to a trusted
LibreOffice executable when automatic executable discovery cannot find it.
Rejected exports remain private staging files; their existence is not approval.
Runtime receipts bind approved paths to exact bytes and are checked again by file
preview, download, and native-open routes. This guards omitted validation and
ignored rejections; it does not sandbox arbitrary host processes.

Delivery checks inspect workbook content as well as its filename, so renaming an
OOXML workbook does not remove the approval requirement. Protected extensions
apply across the owning profile's workspaces to prevent relocation bypasses;
unchanged, explicitly supplied inputs retain their original-byte exemption.
Malformed archive framing and unavailable required classifiers fail closed.
Ordinary chat streams normally, and native directory opens remain available.

Native archive inspection transfers at most 4 MiB of metadata in bounded ranges.
Workbook opens and local workbook downloads use an immutable snapshot capped at
64 MiB; ordinary native archives are copied from the inspected handle without that
size cap. Remote exports use a 64 MiB snapshot cap when the required content policy
is enabled. These are explicit resource limits, not claims that larger files passed
validation.

The public task runner accepts `--enforce-delivery` to apply this same policy to
both arms. It retains blocked attempts and counts a workbook as delivered only
when its final bytes match the runtime receipt. Calculation time is included in
the agent's measured wall time.

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
