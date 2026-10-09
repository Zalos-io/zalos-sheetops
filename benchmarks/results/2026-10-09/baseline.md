# Public spreadsheet benchmark baseline

The final primary score is **20/30 correct for each arm**. SheetOps plus the skill had one timeout versus three for the baseline; each arm produced one file that native Excel rejected. On 17 both-correct pairs, median paired speedup was 1.012x. One repetition per arm does not establish a stable latency advantage.

For a 400,000-row sheet, correct scalar reads near the bottom took 12.06s with OpenPyXL, 5.11s with the SheetOps CLI, and 1.26s with fastexcel. Near the top, OpenPyXL took 0.068s versus 0.102s for the SheetOps CLI. All timings include process startup. Date-containing reads are excluded from SheetOps speed claims because the typed-value contract failed: the raw XML path returns Excel serials. This is a capability gap; preserving the raw-value option matters.

The next changes should add an explicit typed date view, constrain workbook writes to the requested scope, and validate formula serialization and calculated results in the target Excel engine. Four invalid SheetOps calls also give us concrete skill examples to improve. Writes in this experiment used agent-authored Python, so the output failures do not prove a SheetOps write API defect. The two rejected files use modern functions without their future-function prefixes; that is a candidate cause, not a causally isolated finding. [XlsxWriter's formula documentation](https://xlsxwriter.readthedocs.io/working_with_formulas.html) describes the required prefixes and dynamic-array handling.

The native audit covers 87 evaluation copies: 48 formula-containing files were recalculated and saved in Excel 16.113.3, two were rejected without repair, and 37 value-only copies were parsed without native UI calculation. Every source remained intact; all 30 usable golden files pass their self-check. LibreOffice 26.2.4.2 grades are retained separately and differ on formula interoperability.

Run status: complete.

Operations: 480 measurements. Feature contracts: 14 checks.

| Reader | Pass | Incorrect | Unsupported/error |
|---|---:|---:|---:|
| calamine | 63 | 0 | 33 |
| fastexcel | 63 | 0 | 33 |
| openpyxl | 96 | 0 | 0 |
| sheetops-api | 54 | 42 | 0 |
| sheetops-cli | 54 | 42 | 0 |

Reader timing and memory

Cells show median wall time including worker startup / median peak single-process RSS. Timings appear only when every recorded repetition passes. Date-containing ranges use a typed-value contract; scalar ranges exclude the date column.

| Case | calamine | fastexcel | openpyxl | sheetops-api | sheetops-cli |
|---|---:|---:|---:|---:|---:|
| 100000-cells-end | 0.557s / 194.0 MiB | 0.368s / 153.1 MiB | 3.001s / 40.0 MiB | incorrect (3) | incorrect (3) |
| 100000-cells-middle | 0.553s / 193.8 MiB | 0.356s / 152.5 MiB | 1.506s / 35.6 MiB | incorrect (3) | incorrect (3) |
| 100000-cells-start | 0.553s / 194.8 MiB | 0.361s / 152.4 MiB | 0.066s / 31.9 MiB | incorrect (3) | incorrect (3) |
| 100000-scalar-cells-end | 0.569s / 194.8 MiB | 0.362s / 152.2 MiB | 2.987s / 39.9 MiB | 1.327s / 35.2 MiB | 1.347s / 30.9 MiB |
| 100000-scalar-cells-middle | 0.550s / 194.3 MiB | 0.367s / 152.1 MiB | 1.502s / 35.6 MiB | 0.704s / 35.2 MiB | 0.723s / 30.8 MiB |
| 100000-scalar-cells-start | 0.550s / 193.9 MiB | 0.355s / 152.2 MiB | 0.066s / 31.7 MiB | 0.071s / 32.3 MiB | 0.096s / 30.8 MiB |
| 100000-table | 1.172s / 235.1 MiB | 1.132s / 287.1 MiB | 3.601s / 145.3 MiB | incorrect (3) | incorrect (3) |
| 25000-cells-end | 0.187s / 73.7 MiB | 0.161s / 80.0 MiB | 0.789s / 33.6 MiB | incorrect (3) | incorrect (3) |
| 25000-cells-middle | 0.185s / 73.6 MiB | 0.160s / 79.9 MiB | 0.429s / 32.8 MiB | incorrect (3) | incorrect (3) |
| 25000-cells-start | 0.184s / 73.7 MiB | 0.161s / 79.9 MiB | 0.067s / 31.8 MiB | incorrect (3) | incorrect (3) |
| 25000-scalar-cells-end | 0.187s / 73.7 MiB | 0.159s / 79.9 MiB | 0.791s / 33.7 MiB | 0.383s / 35.1 MiB | 0.406s / 30.9 MiB |
| 25000-scalar-cells-middle | 0.184s / 73.5 MiB | 0.158s / 79.9 MiB | 0.422s / 32.8 MiB | 0.224s / 34.9 MiB | 0.251s / 31.0 MiB |
| 25000-scalar-cells-start | 0.186s / 73.6 MiB | 0.160s / 79.9 MiB | 0.065s / 31.7 MiB | 0.071s / 32.4 MiB | 0.098s / 30.9 MiB |
| 25000-table | 0.308s / 77.6 MiB | 0.318s / 107.8 MiB | 0.941s / 59.9 MiB | incorrect (3) | incorrect (3) |
| 400000-cells-end | 2.075s / 679.4 MiB | 1.178s / 441.1 MiB | 11.980s / 67.0 MiB | incorrect (3) | incorrect (3) |
| 400000-cells-middle | 2.090s / 679.4 MiB | 1.201s / 441.1 MiB | 6.128s / 48.7 MiB | incorrect (3) | incorrect (3) |
| 400000-cells-start | 2.068s / 680.6 MiB | 1.176s / 440.6 MiB | 0.067s / 31.6 MiB | incorrect (3) | incorrect (3) |
| 400000-scalar-cells-end | 2.079s / 679.6 MiB | 1.256s / 440.6 MiB | 12.060s / 67.2 MiB | 5.200s / 35.6 MiB | 5.112s / 30.9 MiB |
| 400000-scalar-cells-middle | 2.083s / 680.5 MiB | 1.169s / 440.8 MiB | 5.923s / 48.6 MiB | 2.606s / 35.1 MiB | 2.647s / 30.9 MiB |
| 400000-scalar-cells-start | 2.104s / 680.3 MiB | 1.168s / 440.7 MiB | 0.068s / 31.7 MiB | 0.075s / 32.7 MiB | 0.102s / 31.0 MiB |
| 400000-table | 4.651s / 688.6 MiB | 4.662s / 1003.0 MiB | 14.361s / 491.0 MiB | incorrect (3) | incorrect (3) |
| fidelity-cell_values-cell_values | unsupported (3) | unsupported (3) | 0.065s / 31.4 MiB | incorrect (3) | incorrect (3) |
| fidelity-formulas-References | unsupported (3) | unsupported (3) | 0.067s / 31.2 MiB | 0.065s / 31.7 MiB | 0.091s / 30.9 MiB |
| fidelity-formulas-formulas | unsupported (3) | unsupported (3) | 0.068s / 31.4 MiB | 0.068s / 31.7 MiB | 0.092s / 30.9 MiB |
| fidelity-hyperlinks-Targets | unsupported (3) | unsupported (3) | 0.064s / 31.3 MiB | 0.064s / 31.8 MiB | 0.087s / 30.8 MiB |
| fidelity-hyperlinks-hyperlinks | unsupported (3) | unsupported (3) | 0.069s / 31.2 MiB | 0.063s / 31.6 MiB | 0.094s / 30.9 MiB |
| fidelity-merged_cells-merged_cells | unsupported (3) | unsupported (3) | 0.064s / 31.3 MiB | 0.063s / 31.6 MiB | 0.101s / 30.9 MiB |
| fidelity-multiple_sheets-Alpha | unsupported (3) | unsupported (3) | 0.072s / 31.5 MiB | 0.069s / 31.7 MiB | 0.098s / 30.8 MiB |
| fidelity-multiple_sheets-Beta | unsupported (3) | unsupported (3) | 0.067s / 31.2 MiB | 0.067s / 31.7 MiB | 0.089s / 30.8 MiB |
| fidelity-multiple_sheets-Gamma | unsupported (3) | unsupported (3) | 0.063s / 31.3 MiB | 0.062s / 31.5 MiB | 0.086s / 30.8 MiB |
| fidelity-number_formats-number_formats | unsupported (3) | unsupported (3) | 0.074s / 31.3 MiB | incorrect (3) | incorrect (3) |
| fidelity-tables-tables | unsupported (3) | unsupported (3) | 0.063s / 31.5 MiB | 0.066s / 31.7 MiB | 0.089s / 31.0 MiB |

Paired agent tasks

Completed pairs: 33/33; scorable pairs: 30; both-correct pairs: 17.
Evaluation engines: {"name": "Microsoft Excel", "version": "16.113.3"}.

- baseline: 33 attempts; 30 scorable / 3 unscorable; outcomes {'pass': 20, 'timeout': 3, 'invalid_output': 1, 'incorrect': 6, 'ambiguous_task_metadata': 2, 'oracle_preservation_conflict': 1}; inference cost unavailable.
  Observed SheetOps JSON responses: 0 tasks, 0 successful / 0 failed; commands {}; errors {}.
- sheetops: 33 attempts; 30 scorable / 3 unscorable; outcomes {'incorrect': 8, 'pass': 20, 'timeout': 1, 'invalid_output': 1, 'ambiguous_task_metadata': 2, 'oracle_preservation_conflict': 1}; inference cost unavailable.
  Observed SheetOps JSON responses: 33 tasks, 62 successful / 4 failed; commands {'inspect': 33, 'cells read': 31, 'read-table': 1, 'formulas scan': 1}; errors {'invalid_field': 2, 'invalid_sheet': 1, 'invalid_header_strategy': 1}.

Ranked follow-ups

- P1 Expose date-aware values in bounded XML and table reads: 84 affected attempts. Reproduction IDs: 100000-cells-end, 100000-cells-middle, 100000-cells-start, 100000-table, 25000-cells-end, 25000-cells-middle, 25000-cells-start, 25000-table, 400000-cells-end, 400000-cells-middle, 400000-cells-start, 400000-table, fidelity-cell_values-cell_values, fidelity-number_formats-number_formats.

  Reproduce: `uv run python -m zalos_sheetops.benchmarks.public.runner replay --data DATA --output REPLAY --case 100000-cells-end --adapter sheetops-cli`

- P1 Restrict writes and preserve unrelated cells: 6 affected attempts. Reproduction IDs: 120-24, 35742, 48975, 54638.

  Reproduce: `Run agents.py with the recorded task, model, skill hashes and images`

- P1 Validate XLSX formula serialization in native Excel: 2 affected attempts. Reproduction IDs: 44389, 51680.

  Reproduce: `Run agents.py with the recorded task, model, skill hashes and images`

- P2 Validate answer values and string boundaries: 10 affected attempts. Reproduction IDs: 230-16, 35742, 398-14, 45300, 48975, 51680, 54638.

  Reproduce: `Run agents.py with the recorded task, model, skill hashes and images`

- P2 Apply the requested visible formatting: 6 affected attempts. Reproduction IDs: 343-20, 398-14, 51680, 54638.

  Reproduce: `Run agents.py with the recorded task, model, skill hashes and images`

- P2 Bound workbook work and surface incomplete results: 4 affected attempts. Reproduction IDs: 279-23, 51-12, 547-18.

  Reproduce: `Run agents.py with the recorded task, model, skill hashes and images`

- P2 Preserve workbook structure: 3 affected attempts. Reproduction IDs: 44389, 54638.

  Reproduce: `Run agents.py with the recorded task, model, skill hashes and images`

- P2 Validate formulas with the target Excel engine: 1 affected attempts. Reproduction IDs: 44389.

  Reproduce: `Run agents.py with the recorded task, model, skill hashes and images`

- P3 Resolve reference conflicts before scoring: 6 affected attempts. Reproduction IDs: 168-17, 395-36, 402-43.

  Reproduce: `Run agents.py with the recorded task, model, skill hashes and images`

- P3 Reduce invalid SheetOps calls in the spreadsheet skill: 4 affected attempts. Reproduction IDs: 47827, 52220, 531-18, 57590.

Measurement limits

- Task scores verify the supplied workbook result; future-sheet macro behavior and VBA execution are not tested
- Failure groups overlap: one attempt can fail values, styles and preservation
- Fresh processes, uncontrolled OS file cache; one host; three operation repetitions
- RSS is maximum for one worker or CLI subprocess, not aggregate memory
- Fidelity-prefixed operation cases compare cell values/formulas/caches, not every workbook feature
- Feature contracts separately cover inventory, merges, number formats, hyperlinks and decimal reconciliation
- Agent scores use golden answer-region values, explicit styles and preservation checks; not official leaderboard scores
- Formula evaluation engines are recorded per attempt; native Excel caches are audited independently of raw outputs. LibreOffice results are a separate interoperability check when native results are available
- Native Excel opens formula-containing evaluation copies; value-only outputs are parsed without a native UI open check
- Observed SheetOps commands are inspection and reads; workbook edits use agent-authored Python. Output failures do not establish a SheetOps write API defect
- The first 24 agent attempts lack per-attempt execution source hashes; image, launcher, skill and input hashes remain recorded
- One agent repetition per arm in this initial baseline; insufficient for a stable end-to-end speed claim
- Inference cost is unavailable when provider cost status is unknown, even if reported number is zero

Source attribution

- Mixed-column generator adapted from nazmul284/excel-read-benchmark (MIT), revision d8ddc2b5cfab51d53e196015d7037b1d5f463a8a. https://github.com/nazmul284/excel-read-benchmark/tree/d8ddc2b5cfab51d53e196015d7037b1d5f463a8a
- Selected feature workbooks from SynthGL/ExcelBench (MIT), revision 5f71713413c43f3733e303ba8a4eed6a756760cd. https://github.com/SynthGL/ExcelBench/tree/5f71713413c43f3733e303ba8a4eed6a756760cd
- Selected Verified-400 tasks from KAKA22/SpreadsheetBench (CC-BY-SA-4.0), revision ab0b742b0fc95b946f212d80ac7771b5531272e4. https://huggingface.co/datasets/KAKA22/SpreadsheetBench/tree/ab0b742b0fc95b946f212d80ac7771b5531272e4

These are our measurements and scoped checks, not the upstream projects' published results or leaderboard scores. Source URLs and content hashes are recorded in sources.json; runtime versions, input hashes and code hashes are in run metadata.

Reproduce saved failures

[Download the failure archive](failure-reproductions.zip). It contains the public input, unchanged golden file, original failed output where one exists, task metadata and grading records for all 20 failed scorable attempts. Each output retains its original recorded hash. Original agent logs and native-saved workbooks are excluded. The native grading JSON and audit preserve the evaluated-copy hashes and measured outcomes.

After extracting the archive into REPRO, prepare isolated evaluation copies:

```bash
uv run python -m zalos_sheetops.benchmarks.public.native prepare --dataset REPRO/data --tasks REPRO/tasks.json --runs REPRO/agents --output EVALUATION --excel-version YOUR_EXCEL_VERSION
```

Open each queued copy in Excel, recalculate, save, and close it. Decline any content-recovery prompt and record that observation as `nativeRejection` in its manifest entry. Then verify and grade:

```bash
uv run python -m zalos_sheetops.benchmarks.public.native verify --output EVALUATION
uv run python -m zalos_sheetops.benchmarks.public.regrade --dataset REPRO/data --tasks REPRO/tasks.json --runs REPRO/agents --output REPRO/regraded.jsonl --native-evaluation-root EVALUATION
```

This replays workbook grading without another model run. To repeat inference, use the pinned source tasks, model, image recipes, skill hashes and runner instructions in the repository README. Model sampling and host conditions can change a new inference run. An interrupted local attempt and superseded timing runs remain local and are not counted in this baseline.
