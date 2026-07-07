# Agentic Excel / spreadsheet evals: prior art and recommendations

_Date: 2026-07-07_

This report surveys public spreadsheet/Excel-adjacent benchmarks relevant to a `pi-for-excel` eval suite: an agent running inside real Excel via an Office.js task pane, with structured workbook tools plus an `execute_office_js` escape hatch.

## Executive takeaways

1. **The strongest direct prior art is SpreadsheetBench v1 + SpreadsheetBench 2.** v1 gives real forum-derived manipulation tasks with online-judge-style multi-test-case grading; v2 moves to end-to-end business workflows with hundreds of cell changes and expert validation.
2. **Most older spreadsheet-agent benchmarks under-test real Excel.** SheetCopilot and SheetRM are useful for operation coverage and grading checklists, but SheetCopilot tasks are synthetic/templated and SheetRM’s public release lacks an explicit license.
3. **Finance workflow benchmarks are converging on rubric + perturbation grading.** MBABench, BlueFin, and Finch show why exact cell matching is insufficient for large professional workbooks: dynamic formulas, auditability, format conventions, and scenario perturbations matter.
4. **For pi-for-excel, keep real Excel as the primary engine.** LibreOffice/openpyxl ports are useful for scale, but the product target is Excel; many benchmarks explicitly run into formula/function fidelity gaps outside Excel.
5. **Use deterministic final-state grading wherever possible.** LLM/VLM judges are useful for charts and professional deliverables, but they should not be the regression gate for core tool-surface correctness.
6. **Legally reusable corpuses are mixed.** SpreadsheetBench/SpreadsheetBench 2 are CC BY-SA 4.0; Finch is CC BY 3.0; TableBench data is Apache-2.0; InstructExcel is MIT but points to third-party workbook URLs. SheetCopilot and InfiAgent data are non-commercial; SheetRM/MiMoTable/MMTU/BlueFin licensing needs confirmation before reuse.

---

## Comparison table

| Benchmark / work | What it tests | Size + provenance | Task spec | Grading | Harness / engine | Metrics | Openness / license | Caveats for pi-for-excel |
|---|---|---:|---|---|---|---|---|---|
| [SpreadsheetBench v1](https://github.com/RUCKBReasoning/SpreadsheetBench), [paper](https://arxiv.org/abs/2406.14991), [site](https://spreadsheetbench.github.io/) | Real-world spreadsheet manipulation, mostly answer-producing edits | 912 instructions; 2,729 test spreadsheets; real questions from online Excel forums/blogs | JSONL with instruction, spreadsheet path, instruction type, `answer_position` | OJ-style final-state exact match over multiple similar input/output test cases; formulas recalculated first | LLM generates code; code executed in Docker; eval with openpyxl after LibreOffice or Excel/win32com recalculation | Pass@1 across all spreadsheets; cell/sheet-level categories | CC BY-SA 4.0; full 912 released; Verified subset of 400 announced | Excellent realism; answer positions reduce placement ambiguity; authors note corner cases are not exhaustively devised |
| [SpreadsheetBench 2](https://spreadsheetbench.github.io/), [paper](https://arxiv.org/abs/2606.29955), [HF](https://huggingface.co/datasets/KAKA22/SpreadsheetBench-v2) | Workflow-level business spreadsheet agents: generation, debugging, visualization | 321 tasks: 100 financial modeling, 100 debugging, 97 template, 24 visualization; authentic business data, financial reports, corporate filings; expert-built/validated | Natural-language workflow instruction + in-progress workbook | For cell tasks: compare target-cell computed values and full-workbook exact match. For charts: expert rubric + VLM judge | SWE-agent-style CLI scaffold with `bash`, `view_xlsx`, `submit`; spreadsheet products manually tested on 30 examples | `Modif.` target-cell match; `Acc.` all-cells task success; visualization pass rate | Dataset CC BY-SA 4.0; eval code MIT | Best model only 34.89% Acc.; debugging 12%. Exact match may undercount alternative valid formulas/layouts; VLM judge adds noise |
| [SheetCopilot Benchmark](https://github.com/BraveGroup/SheetCopilot), [paper](https://arxiv.org/abs/2305.19308) | Agentic spreadsheet control through atomic actions | 221 tasks over 28 workbooks; derived from 67 seed tasks; 44 operations across manipulation, management, formatting, charts, pivots | `dataset.xlsx` lists workbook, context, instruction, categories, atomic actions; reference workbooks + YAML checklists | Final workbook compared with any reference solution; property checklists for cells/charts/pivots/filters/CF/view; numeric tolerance 1e-8 | Original Windows Excel + pywin32; 2026 Ubuntu evaluator uses openpyxl + headless LibreOffice UNO | Exec@1, Pass@1, A_mean/A50/A90 | Repo LICENSE GPLv3, but README says agent/dataset are non-commercial; treat as non-commercial/ambiguous | Great operation taxonomy and checklist examples; synthetic tasks and non-commercial wording make it weaker as a directly reusable corpus |
| [SheetAgent / SheetRM](https://github.com/cybisolated/SheetAgent), [paper](https://arxiv.org/abs/2403.03636) | Spreadsheet reasoning/manipulation as subtasks; planning + API execution | Paper: 137 sheets, 317 tasks, 1,625 subtasks. Public repo subset inspected: 25 spreadsheets, 180 tasks | `tasks.xlsx` with spreadsheet file, context, instruction, manipulation category, reasoning challenge | Paper uses task checklists and Python evaluative criteria for subtask/procedure evaluation | Agent with spreadsheet-specific APIs; public release contains workbooks/tasks, not a polished universal harness | Task/subtask pass rates in paper | No license found in repo; source data from public examination question bank; anonymized | Useful for checklist design and subtask decomposition; do not reuse files without permission/explicit license |
| [OSWorld](https://github.com/xlang-ai/OSWorld), [paper](https://arxiv.org/abs/2404.07972) | General desktop GUI agent benchmark, including LibreOffice Calc | 369 total tasks in `evaluation_examples/test_all.json`; 47 LibreOffice Calc tasks counted locally | JSON configs with instruction, downloaded files, app start/open actions, evaluator function | Final-state evaluators such as table comparison; trajectories logged | Ubuntu VM + accessibility/screenshot control; LibreOffice apps | Task success rate | Apache-2.0 | Valuable harness pattern and permissive tasks, but not Excel/Office.js; Calc fidelity gap |
| [WindowsAgentArena](https://github.com/microsoft/WindowsAgentArena), [paper](https://arxiv.org/abs/2409.08264) | Scalable Windows GUI agents; includes Calc/office-like tasks | 154 tasks in `test_all.json`; 24 `libreoffice_calc` in test set, 48 JSON files present including variants | Windows VM task configs; app activation, pyautogui save steps, final-state evaluator | Final-state task success | Windows VM/Azure parallelization; UI automation | Task success rate | MIT | Useful Windows harness inspiration; Calc tasks are not real Excel; setup heavier than needed for Office.js add-in evals |
| [Spreadsheet-RL / Domain-Spreadsheet](https://arxiv.org/abs/2506.16174) | RL-trained agents for real Excel task automation | Domain-Spreadsheet benchmark: 1,660 tasks. Training corpus: 18,855 threads from Microsoft Answers/TechCommunity, filtered to 5,928 high-quality tasks | Labeled operation categories and expected workbook outcomes | Microsoft Excel 365 execution; final-state checks reported in paper | Excel 365 + UI/API agent | Task pass rates; improves from 12.34% to 20.47% | Paper says benchmark/code/model will be released; license not verified | Important because it uses real Excel and public support threads; not currently reusable until release/license verified |
| [InstructExcel](https://github.com/microsoft/InstructExcel), [paper](https://arxiv.org/abs/2310.14495) | NL-to-OfficeScript code generation | Repo JSON inspected: 10,520 examples. Paper experiments: filtered splits of 4,033 train / 1,000 dev / 1,000 test, plus 200-sample test subset | Prompt + workbook URL + operation metadata + OfficeScript solution | Code generation judged by similarity/execution in paper; not a real agent final-state benchmark | OfficeScript / Excel APIs | EM / similarity / execution metrics | MIT repo license; some workbook URLs are missing/broken; third-party workbook provenance | Excellent source of Office.js/OfficeScript idioms; less useful for measuring end-to-end interactive agents |
| [InfiAgent-DABench / DAEval](https://github.com/InfiAgent/InfiAgent) | Data-analysis answers over CSV tables, not workbook manipulation | Paper: 257 questions over 52 CSV files. Public validation README says 400 questions over 72 CSV files | Question + CSV(s); answer required in `@name[...]` format | Closed-form exact/numeric check; numeric tolerance `< 1e-6` | Agent generates analysis/code over CSV files | Accuracy | Code Apache-2.0; data CC BY-NC 4.0 | Useful for analytics-answer grading patterns; non-commercial data and not Excel state manipulation |
| [TableBench](https://github.com/TableBench/TableBench), [paper](https://arxiv.org/abs/2407.06437) | Complex table QA/reasoning and spreadsheet-like operations | 886 human-written test cases; 18 fine-grained task categories, 4 broad categories | QA over tabular data; benchmark and scoring scripts | Answer correctness, split by task/category | LLM/table QA harness | Accuracy | Code MIT; data Apache-2.0 per repo badges | Permissive adjacent reasoning corpus; no workbook object model, formatting, charts, or final-state edits |
| [MBABench](https://github.com/namkoong-lab/MBABench), [paper](https://arxiv.org/abs/2505.16225) | MBA/analyst-style financial spreadsheet modeling tasks | End-to-end financial workbooks from FMWC, ModelOff, Wall Street Prep; exact full task count not clearly stated in accessible text inspected; public HF subset `mbabench-modeloff` appears to have 38 rows | Instruction + financial workbook; expected analyst deliverable | LLM/judge rubric over Accuracy, Formula, Format, with fine-grained criteria | Spreadsheet agents / product comparisons; evaluation reportedly expensive (~$1.7K in considered setting) | Rubric score | Repo MIT license; source workbook rights may vary by underlying competitions/providers | Very relevant for finance-quality grading; reuse requires care around source data rights and LLM judge cost/noise |
| [BlueFin](https://arxiv.org/abs/2505.06192) | Financial spreadsheet reasoning and modification | 131 tasks; 3,225 granular rubric criteria; synthesis, manipulation, comprehension | Real-world analyst workflows and public-company financial data | Expert-validated LM judge; reported human agreement alpha 0.826 / macro-F1 0.839 | Open-source harness claimed in paper, but repo/license not verified | Rubric scores; strongest LLMs under 50% average | License/repo not verified | Strong rubric-design reference; do not reuse until license and artifact URL are verified |
| [Finch / FinWorkBench](https://huggingface.co/datasets/finosfoundation/finworkbench-finch) | Enterprise finance/accounting workflows across spreadsheets + PDFs/images/code/web | 172 workflows; 1,710 spreadsheets; 27M+ cells; average 8 sheets/workbook | Multi-file workflow instructions; derived from Enron, EUSES, recent public filings/artifacts | Paper describes perturbation robustness and granular workflow evaluation | Agentic finance workflow harness | Workflow success / component metrics | HF dataset license CC BY 3.0 | Highly relevant for realistic finance contexts; broader than Excel and may require multi-file retrieval/internet tools |
| [MiMoTable](https://arxiv.org/abs/2408.09102) | Multi-scale spreadsheet QA and operations | 428 real-world spreadsheets across 7 domains; 1,719 QA pairs | Questions covering lookup, edit, compare, calculate, visualize, reasoning | QA/op correctness | Table/spreadsheet benchmark harness | Accuracy by category | License not verified | Useful for operation taxonomy and multi-scale sheets; not safe to reuse until license found |
| [MMTU](https://github.com/MMTU-Benchmark/MMTU), [paper](https://arxiv.org/abs/2401.15385) | Massive multi-task table understanding/reasoning/manipulation | 28,136 questions; 25 tasks; 46 datasets per README | QA tasks over many table datasets | Answer scoring | Table benchmark | Accuracy | License not found in repo/HF during inspection | Broad table benchmark, not Excel-agent eval; license ambiguity |
| [SpreadsheetLLM](https://arxiv.org/abs/2407.09025) | Spreadsheet representation/encoding for LLM understanding | Evaluates table detection and QA using compressed sheet encodings | SheetCompressor / SheetEncoder formats | Table detection F1 and QA metrics | LLM encoding, not workbook automation | 25.6% GPT-4 table-detection gain; fine-tuned model reaches 78.9% F1 with ~25x compression | Paper / project references; license not material for corpus reuse here | Useful for observation design (`view_xlsx`/range summarization), not a task corpus |
| [FLARE](https://arxiv.org/abs/2505.07505) | Formula Logic and Reasoning Evaluation; spreadsheet error detection/repair | Modular EuSpRIG 2025 benchmark; selected error-seeded spreadsheet tasks, e.g. Computer Chip Factory and Triangle | Prompts + spreadsheets; full prompts/responses via Google Drive in paper | Human-assessed formula reasoning and error-detection outcomes | LLM prompt benchmark | Task scores | Paper artifacts available, but explicit reusable dataset license not verified | Good source of formula-error patterns; probably use as inspiration unless license clarified |

---

## Per-benchmark notes

### 1. SpreadsheetBench v1

**Why it matters.** This is the cleanest public example of realistic spreadsheet-manipulation grading. It was built from real Excel forum/blog questions rather than self-instruct templates, and each instruction has multiple structurally similar test cases with different data values. That makes it closer to software OJ evaluation than single-workbook “copy the reference file” grading.

**Facts to carry forward.**

- 912 instructions and 2,729 test cases, average roughly 3 test cases per instruction.
- Spreadsheets include multiple tables, non-standard relational tables, and non-textual elements such as color/bold formatting.
- Construction used human annotation; the paper says an annotation team of 20 Excel specialists was used.
- The benchmark is CC BY-SA 4.0 and the authors explicitly discuss copyright risk from forum-derived data; they say they avoid redistributing raw source data.
- The authors disclose limitations: posts without acknowledged responses or hard-to-formalize questions were filtered out, and they did not meticulously devise every possible corner case for every question.

**Design lesson.** OJ-style data mutation is a strong antidote to overfitting exact visible values. For pi-for-excel, we should copy the pattern, not necessarily the files: create seed workbooks + generated variants + hidden oracle assertions.

### 2. SpreadsheetBench 2

**Why it matters.** SpreadsheetBench 2 shifts from isolated formula/formatting edits to full business workflows. Its tasks average 11.8 worksheets and 593.5 cell modifications per instance, which is far closer to “make this workbook useful” than one-cell answer benchmarks.

**Facts to carry forward.**

- 321 tasks: 100 financial modeling, 100 debugging, 97 template generation, 24 visualization.
- Authentic business data: annual reports, corporate filings, and professional spreadsheet artifacts.
- Grading has two cell-task levels:
  - `Modif.`: target-cell computed values match expected outputs.
  - `Acc.`: all relevant workbook cells match, i.e. end-to-end task success.
- Visualization tasks use a rubric and VLM judge.
- Reported best overall task accuracy is 34.89%; debugging accuracy is only 12.00%.
- Spreadsheet products were manually tested on a 30-example subset; none surpassed the agent scaffold, and Claude for Excel was reported as best spreadsheet product at 15.4% on that subset.

**Design lesson.** It is useful to track both partial modification correctness and full task success. For pi-for-excel, a `modification_score` can aid debugging, but release gates should use task-level pass/fail for the user-visible workflow.

### 3. SheetCopilot

**Why it matters.** SheetCopilot is still the best operation-taxonomy benchmark for spreadsheet agents: 44 operations, including formatting, charts, pivot tables, conditional formatting, filters, and view state.

**Facts to carry forward.**

- 221 tasks over 28 spreadsheets.
- Original evaluator uses Windows Excel 2019 and pywin32; a newer Ubuntu evaluator in the repo uses openpyxl plus headless LibreOffice UNO and claims the same outcome metrics.
- Metrics: `Exec@1`, `Pass@1`, and action-level scores `A_mean`, `A50`, `A90`.
- Repo license is GPLv3, but README text says the dataset/agent are for non-commercial research/education. This is inconsistent enough that commercial reuse should be avoided without permission.

**Design lesson.** Reuse the checklist idea. Its evaluator decomposes workbook objects into checkable properties rather than only comparing files byte-for-byte; that is essential for charts, pivots, filters, frozen panes, conditional formats, and formatting.

### 4. SheetAgent / SheetRM

**Why it matters.** SheetRM emphasizes decomposition into tasks/subtasks and grading via checklists plus Python evaluative criteria. That is closer to an agent trace/evaluation design than pure final answer QA.

**Facts to carry forward.**

- Paper describes full SheetRM as 137 sheets, average 300.82 rows and 26.23 columns, 317 tasks, and 1,625 subtasks.
- The public repo subset inspected contains 25 spreadsheets and 180 tasks under `sheetrm`.
- `tasks.xlsx` columns: `Spreadsheet File`, `Context`, `Instruction`, `Manipulation Category`, `Reasoning Challenge`.
- The paper says the dataset was built from a public examination question bank and privacy-modified/anonymized.
- No explicit license file was found in the repo.

**Design lesson.** Use subtasks for diagnostic reporting, but do not overfit grading to an expected trajectory. The pi-for-excel harness should log tool traces and optionally score subtask evidence, while the authoritative result remains the workbook final state.

### 5. OSWorld and WindowsAgentArena

**Why they matter.** These are not Excel benchmarks, but they are the closest mature examples of desktop-agent final-state harnessing.

**Useful pieces.**

- OSWorld: Apache-2.0; 369 tasks in the main evaluation file, 47 LibreOffice Calc tasks counted locally.
- WindowsAgentArena: MIT; 154 tasks in `test_all.json`, 24 `libreoffice_calc` tasks there, with more variants present in the repo.
- Both represent tasks as JSON specs with setup, app activation/download/open steps, and final-state evaluator functions.

**Design lesson.** The task spec should be declarative and runnable from a clean VM/session. But pi-for-excel should avoid unnecessary screenshot/desktop-control complexity for the core lane; an Office.js add-in can run a much thinner, more deterministic harness inside Excel.

### 6. Spreadsheet-RL / Domain-Spreadsheet

**Why it matters.** Spreadsheet-RL is notable because it uses Microsoft Excel 365 as the environment, not a simulator. It also mines real user support threads and trains agents with verifiable outcomes.

**Facts to carry forward.**

- Domain-Spreadsheet benchmark: 1,660 tasks.
- Training corpus starts from 18,855 Microsoft Answers / TechCommunity threads, filtered to 5,928 high-quality tasks.
- Paper reports performance improvement from 12.34% to 20.47% on the benchmark.
- Release and license were not verified; treat as non-reusable until artifacts are actually available.

**Design lesson.** This validates the local `pi-for-excel` philosophy: real Excel primary lane, no simulator as source of truth.

### 7. InstructExcel

**Why it matters.** InstructExcel is not an interactive agent benchmark, but it is highly relevant to the Office.js/OfficeScript tool surface.

**Facts to carry forward.**

- Repo JSON inspected contains 10,520 examples.
- Paper reports filtered experimental splits of 4,033 train / 1,000 dev / 1,000 test, and a 200-sample test subset due to GPT cost.
- License is MIT, but some workbook URLs are missing/broken and many workbooks originate from external sources.

**Design lesson.** Use it for OfficeScript idioms, function/tool coverage, and synthetic task inspiration. Do not treat OfficeScript code similarity as a sufficient measure of an agent that must inspect, reason, edit, and recover inside a live workbook.

### 8. Data/table reasoning benchmarks: DABench, TableBench, MMTU, MiMoTable, SpreadsheetLLM

These are adjacent, not direct substitutes.

- **DABench / InfiAgent.** Good closed-form answer extraction pattern. Version discrepancy must be preserved: the paper says 257 questions over 52 CSV files; repo README for public validation says 400 questions over 72 CSV files. Data is CC BY-NC 4.0.
- **TableBench.** Permissive table reasoning dataset: code MIT and data Apache-2.0. Useful for reasoning cases, but lacks workbook object state.
- **MMTU.** Very broad table benchmark: 28,136 questions, 25 tasks, 46 datasets. License not found during inspection.
- **MiMoTable.** 428 spreadsheets, 1,719 QA pairs, six meta-operation categories: Lookup, Edit, Compare, Calculate, Visualize, Reasoning. License not verified.
- **SpreadsheetLLM.** Useful for observation/compression design: SheetCompressor/SheetEncoder improves table detection and compresses large sheets, but it is not a manipulation benchmark.

### 9. Financial workflow benchmarks: MBABench, BlueFin, Finch

**Why they matter.** These are the right inspiration for “would a finance professional trust this workbook?” rather than “did cell B7 equal 42?”. They evaluate formulas, formatting, business logic, and robustness.

**Notes.**

- **MBABench.** End-to-end financial spreadsheet tasks from FMWC, ModelOff, and Wall Street Prep sources. Rubric dimensions are Accuracy, Formula, and Format. Repo license is MIT, but underlying workbook rights and exact full task count need clarification before reuse.
- **BlueFin.** 131 financial spreadsheet tasks and 3,225 granular rubric criteria. Paper reports expert validation of an LM judge with alpha 0.826 and macro-F1 0.839. Exact repo/license not verified.
- **Finch / FinWorkBench.** 172 workflows, 1,710 spreadsheets, 27M+ cells, average 8 sheets per workbook; HF dataset license CC BY 3.0. It spans spreadsheets, PDFs, images, code, web/search, retrieval, modeling, validation, visualization, and reporting.

**Design lesson.** For a professional Excel eval lane, add formula correctness, formatting/readability, and perturbation tests. Exact cell values alone can reward hardcoding and punish dynamic models incorrectly.

### 10. Public product benchmark claims

These claims are useful context but should not drive local scoring without reproduction.

- Microsoft reports Excel Agent Mode scoring **57.2%** on the full 912-task SpreadsheetBench v1. The blog says Agent Mode uses Excel APIs in a JavaScript runtime and that accuracy was measured using the SpreadsheetBench authors’ openpyxl grading script. Treat as vendor-reported unless independently reproduced.
- OpenAI reports ChatGPT agent scoring **45.54%** on SpreadsheetBench with `.xlsx` files and reports Microsoft Copilot at **20%** in its comparison. This appears to use a different harness/environment from Microsoft’s Excel-native claim.
- SpreadsheetBench 2 reports product tests on only a 30-example subset, manually run by humans; useful as a directional signal, not an apples-to-apples benchmark.

---

## Reusable corpuses and legal risk

### Relatively safe candidates, subject to attribution/share-alike fit

1. **OSWorld** — Apache-2.0. Good for harness/evaluator examples and permissive Calc tasks.
2. **WindowsAgentArena** — MIT. Good for Windows VM task spec examples; less directly useful for Excel.
3. **TableBench** — code MIT, data Apache-2.0. Good for table-reasoning seed ideas.
4. **InstructExcel** — MIT repo. Good for OfficeScript/Office.js idioms, but validate external workbook URLs and third-party source rights before copying workbook files.
5. **Finch / FinWorkBench** — CC BY 3.0. Strong finance workflow context if attribution and data handling fit.

### Reusable only if share-alike is acceptable

1. **SpreadsheetBench v1** — CC BY-SA 4.0. Strong source for realistic tasks; share-alike may be awkward if incorporated into a proprietary benchmark corpus. Safer pattern: use it as a baseline benchmark run or inspiration, not as mixed-in proprietary training/eval data.
2. **SpreadsheetBench 2** — dataset CC BY-SA 4.0; eval code MIT. Same share-alike caution.

### Avoid commercial reuse without permission / clarification

1. **SheetCopilot** — GPLv3 file plus README non-commercial wording. Treat as non-commercial/ambiguous.
2. **InfiAgent-DABench data** — CC BY-NC 4.0.
3. **SheetAgent / SheetRM** — no explicit license found; do not reuse files.
4. **BlueFin** — repo/license not verified.
5. **MiMoTable** — license not verified.
6. **MMTU** — license not found during inspection.
7. **FLARE artifacts** — explicit reusable data license not verified.
8. **MBABench underlying workbooks** — repo MIT, but competition/provider source rights need checking before copying workbooks.

### Practical legal recommendation

For an internal pi-for-excel regression suite, prefer **freshly authored seed workbooks and task specs**. Use public benchmarks in three bounded ways:

- Run external benchmarks unmodified as third-party baselines where their licenses permit.
- Study taxonomies, schemas, and grading ideas.
- Recreate analogous tasks using newly authored or clearly licensed workbooks, not copied proprietary/community artifacts.

---

## Recommendations for the pi-for-excel eval suite

These recommendations align with the local eval proposal: real Excel is the primary lane; keep the harness thin; use durable task specs, seed workbooks, and grading assertions; do not build a simulator.

### 1. Split the suite into three lanes

**Lane A — atomic tool regression.** 30-80 small tasks covering the Office.js tool surface: read/write, formulas, fills, formats, tables, filters, sort, conditional formatting, names, sheet operations, charts, pivots, data validation, comments/notes if supported. These should be deterministic, fast, and mostly fixed-output-location.

**Lane B — realistic manipulation tasks.** 50-150 forum-style and business-admin tasks with data variants. Use SpreadsheetBench’s OJ idea: the task prompt is stable, but values/rows/edge cases vary across hidden workbook instances. Grade final state, not code style.

**Lane C — workflow/professional tasks.** A smaller suite, maybe 10-30 tasks, for multi-sheet finance/reporting/debugging/visualization workflows. Track partial metrics, but gate on task-level pass/fail plus human-readable diff reports. This lane can include optional human/LLM review for non-deterministic aesthetics, but the default release gate should remain deterministic assertions.

### 2. Use a durable task spec

Each eval should be a directory or JSON/YAML spec with:

```yaml
id: excel-format-sales-summary-001
title: Format and summarize regional sales
seed_workbook: seed.xlsx
prompt: |
  Add a summary table by region and create a bar chart of total revenue.
allowed_tools: default
setup:
  recalc: true
  freeze_time: 2026-07-07
variants:
  generator: generate_variants.py
  hidden_cases: 3
grading:
  assertions:
    - type: table_exists
      sheet: Summary
      locator: label:Region
    - type: cell_values_equal
      locator: table_column:Summary!Revenue
      expected_from: oracle.py
      tolerance: 0.01
    - type: chart_series
      sheet: Summary
      chart_title_contains: Revenue by Region
      expected_series_count: 1
    - type: no_clobber
      range: Data!A1:G500
telemetry:
  collect_tool_calls: true
  collect_workbook_diff: true
```

The key is that the task spec owns the seed workbook and assertions. The model never sees the oracle, hidden variants, or expected output workbook.

### 3. Grade final workbook state first

Recommended assertion families:

- **Values:** exact strings, numeric tolerance, dates with locale/timezone normalization, booleans/errors.
- **Formulas:** formula text where required; formula result where equivalent formulas are acceptable; dependency/range references for dynamic correctness.
- **Formats:** number formats, fill/font/border/alignment, conditional formatting rules.
- **Objects:** table names/ranges/styles/totals, chart type/title/series/categories, pivot cache/source/fields/filters, named ranges, data validation, filters/sort state, frozen panes.
- **Workbook structure:** sheet existence/order/visibility, protected ranges if used, no unwanted sheets, no clobber outside allowed ranges.
- **Robustness:** rerun recalc; perturb source data; ensure formulas update and hardcoded values fail.
- **Semantic locators:** find by sheet name, table name, label, named range, nearby text, or object metadata. Avoid relying only on A1 coordinates unless the instruction explicitly mandates a location.

### 4. Keep trajectory logs diagnostic, not authoritative

Capture tool calls, token counts, errors, retries, Office.js escape-hatch usage, and workbook diffs. Use them to explain failures and compare tool interfaces. Do **not** fail a task merely because the agent used an unexpected path if the final workbook is correct and safe.

Exceptions worth tracking as policy metrics:

- use of `execute_office_js` when a structured tool exists;
- excessive full-workbook reads;
- destructive edits outside intended ranges;
- inability to recover after Excel/Office.js errors.

### 5. Avoid solution leakage

- Do not put answers in hidden sheets, comments, custom properties, names, external links, or stale cached values the agent can read.
- Generate hidden variants after task authoring, not from the same visible answer workbook.
- Strip workbook metadata, calc chains, comments, and unused hidden ranges from seed files.
- Keep oracle scripts outside the working directory exposed to the agent.
- Do not tune prompts/tools on public benchmark test tasks and then report those tasks as held-out.

### 6. Use data variants to defeat hardcoding

For most Lane B/C tasks, create at least 3 hidden cases:

1. normal case;
2. edge case: blanks, duplicates, negative values, zero denominators, missing categories;
3. scale case: more rows/sheets than the visible example.

SpreadsheetBench’s biggest contribution is not just “real forum questions”; it is the OJ-style insistence that a solution generalize to similarly structured but different workbooks.

### 7. Make placement ambiguity explicit

If the user instruction says “put the answer in H2,” grade H2. If it says “add a summary,” either:

- require a location in the eval prompt, or
- grade semantically by finding a labeled table/chart/output block.

Do not silently require a hidden coordinate for a task whose prompt allowed a different professional layout.

### 8. Add professional finance/readability checks only where needed

For finance-style tasks, borrow MBABench/BlueFin dimensions but implement deterministically where possible:

- **Accuracy:** computed values and scenario outputs match oracle.
- **Formula:** outputs are formulas referencing source drivers, not hardcoded constants; formulas survive perturbation.
- **Format:** currency/percent/date formats, totals/subtotals, consistent signs, headers, freeze panes, print/readability if relevant.
- **Auditability:** named assumptions, clear labels, no hidden broken links, no circular references unless intended.

Use LLM judges only for “would an analyst consider this presentation acceptable?” secondary review, not for the first-pass CI gate.

### 9. Measure both correctness and product ergonomics

Minimum metric set:

- `task_pass`: all required assertions pass.
- `assertion_pass_rate`: diagnostic partial score.
- `destructive_edit_count`: cells/objects changed outside allowed diff.
- `recalc_clean`: no Excel errors/circular refs/unresolved external links.
- `tool_calls`, `elapsed_sec`, `model_tokens`, `excel_errors`.
- `escape_hatch_rate`: fraction of tasks using raw Office.js.
- `variant_pass_rate`: pass rate across hidden data variants.

For tool-interface experiments, keep the benchmark fixed and compare tool surfaces/model providers using the same task specs.

### 10. Start with a small but end-to-end v0

A practical first milestone:

- 10 atomic tasks: formulas, formats, tables, filters, charts.
- 10 realistic manipulation tasks with 2 hidden variants each.
- 3 workflow tasks: one financial model, one debugging workbook, one visualization/report.
- Run in real Excel from a fresh workbook copy.
- Produce a markdown/JSON report with workbook diffs and failed assertions.

This gives enough coverage to validate the harness, grading, and pi-for-excel tool ergonomics before spending time authoring a large corpus.

---

## Recommended source priority

1. **Implement from first principles:** real Excel harness, final-state assertions, workbook diffing, hidden variants.
2. **Copy design patterns, not data, from:** SheetCopilot checklists, SpreadsheetBench OJ-style variants, SpreadsheetBench 2 `Modif.` vs `Acc.`, OSWorld/WAA JSON task specs.
3. **Run as external baselines if useful:** SpreadsheetBench v1/v2, respecting CC BY-SA terms and separation from proprietary corpus.
4. **Use as adjacent inspiration:** InstructExcel for OfficeScript APIs; TableBench/DABench for table reasoning; MBABench/BlueFin/Finch for finance rubrics.
5. **Avoid direct reuse until clarified:** SheetRM, MiMoTable, MMTU, BlueFin artifacts, and any competition/provider workbooks with uncertain redistribution rights.

## Bottom line

The right eval suite for pi-for-excel should not be a simulator, a code-generation benchmark, or a screenshot-only desktop task. It should be a **real Excel final-state benchmark**: durable seed workbooks, natural user prompts, hidden data variants, deterministic workbook assertions, and diagnostic trajectory logs. Public benchmarks give strong design patterns, but the safest and most product-relevant corpus is one we author ourselves, using permissively licensed public data only where the legal path is clear.
