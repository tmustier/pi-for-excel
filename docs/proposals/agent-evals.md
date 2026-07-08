# Proposal: Agent Eval Suite

**Status:** Proposal (not yet accepted) — v2, restructured task-set-first
**Date:** 2026-07-07
**Companion to:** [`agent-tool-interface-redesign.md`](./agent-tool-interface-redesign.md)
**Prior art survey:** [`../research/spreadsheet-agent-evals-prior-art.md`](../research/spreadsheet-agent-evals-prior-art.md)

Telemetry is off the table (product decision) and no usage corpus exists.
Evals are the evidence loop for the tool-surface redesign, the regression gate
for prompt/context/tool changes, and the only honest answer to "does this work
on smaller models?" (#603).

## 1. Design principles

1. **The task set is the asset; the harness stays thin.** Task specs, seed
   workbooks, and grading assertions are durable data — they outlive any
   harness, model, or tool surface. Effort goes there first.
2. **Real Excel is the primary lane, not the fallback.** The
   background-verification bridge already drives the real prompt → model →
   tool loop (`submitPrompt`) and reads real workbook state back
   (`readRange`, `listCharts`, …). The real engine computes formulas, so
   grading on computed values is free. At eval scale (≤50 tasks × minutes
   each, run per-phase rather than per-commit), speed is a non-issue and
   fidelity is everything.
3. **No workbook simulator.** An earlier draft proposed an in-memory fake
   workbook lane. Rejected on cost/fidelity grounds:
   - Formula engines don't solve the problem. HyperFormula (GPLv3/commercial
     dual-licensed) would be license-acceptable as a dev-only harness
     dependency — GPL obligations trigger on distribution, and evals never
     ship in the MIT add-in bundle — but it is a formula engine, not an
     Office.js host emulator. The bulk of the simulator is the host object
     model (formats, charts, pivots, tables, freeze panes, insert/delete
     semantics), which we would still hand-build around it, and its formula
     semantics diverge from real Excel at the edges, so even the part it
     covers isn't ground truth.
   - The real-Excel lane already exists and *is* ground truth, so the
     simulator adds cost without adding trust.

   Unit tests keep using small ad-hoc fakes; evals use real Excel. Revisit
   only if we someday need CI-scale or cross-platform runs.
4. **Deterministic grading.** Assert final workbook state: cell values,
   formula text, number formats, object existence (tables/charts/pivots/
   names), freeze panes, and absence of leftover artifacts. Reply-text checks
   are `contains`-style and coarse. LLM-judge only ever as a labeled
   spot-check, never the gate.
5. **Hidden data variants defeat hardcoding.** SpreadsheetBench's core
   insight: a solution must generalize to structurally similar workbooks with
   different values. For manipulation/workflow tasks, generate 2–3 hidden
   variants from the fixture builder (normal / edge case / scale case) and
   grade across all of them. Cheaper and stronger than formula-text sniffing.
6. **Placement ambiguity is explicit.** If the prompt pins a location, grade
   that location. If it says "add a summary", grade semantically (find the
   labeled block/table) — never silently require a hidden coordinate for a
   task whose prompt allowed a different professional layout.

## 2. Metrics (per task × model)

| Metric | Why |
|---|---|
| **Task success** (assertions pass) | The material outcome |
| **Escape-hatch rate** (`execute_office_js` calls) | Which structured tool to build next — replaces telemetry |
| **Tool-error rate** (failed calls, schema rejections, retries) | Schema/semantics friction |
| **Efficiency** (tool calls, tokens, wall time) | Batch/high-level tool gaps |
| **Behavioral checks** (asked-when-ambiguous, no-clobber-without-confirm) | Interaction quality |
| **Variant pass rate** (pass across hidden data variants) | Anti-hardcoding signal |
| **Destructive-edit count** (cells/objects changed outside allowed diff) | Mutation safety / collateral damage |

Report two levels per task (SpreadsheetBench 2's `Modif.` vs `Acc.` split):
**partial score** (fraction of assertions passing — diagnostic) and **task
pass** (all required assertions — the gate). A near-miss with one convention
error should be visible as such, not as total failure.

Per-tool-surface phase, the acceptance criterion for a new tool is a measured
drop in escape-hatch/failure rate in its category.

## 3. Task set v0 (~20 tasks)

The suite grows along three lanes (per the prior-art survey): **Lane A** —
atomic tool regression (fast, deterministic, fixed locations; most of v0);
**Lane B** — realistic manipulation tasks with hidden variants; **Lane C** —
multi-sheet workflow/professional tasks (model debugging, scenario edits,
build-outs; mostly corpus-derived). Lane C is where current frontier agents
are weakest — SpreadsheetBench 2 reports 12% on debugging for the best
model — and where a real-Excel agent can differentiate.

Seed workbooks are small fixtures (`evals/fixtures/*.xlsx` or builder
scripts); task specs are data (YAML/JSON), e.g.:

```yaml
id: tables-01
category: tables
seed: fixtures/sales-raw.xlsx
prompt: "Turn the data on Sheet1 into a table and sort it by Revenue, highest first."
assertions:
  - table_exists: { sheet: Sheet1 }
  - sorted_by: { column: Revenue, order: desc }
budget: { max_tool_calls: 8 }
```

| ID | Category | Task sketch | Key assertions |
|---|---|---|---|
| orient-01 | Orientation | "What's in this workbook?" (3-sheet model) | No mutation; reply names all sheets |
| orient-02 | Orientation | "Where is FY25 gross margin?" | No mutation; reply cites correct cell |
| formula-01 | Formulas | Add a Total row summing month columns | `=SUM(...)` formula text + computed values |
| formula-02 | Formulas | Add a YoY growth column | Formulas reference prior-year cells; values correct |
| formula-03 | Formulas | Explain a nested-IF cell | No mutation; reply names the input cells |
| formula-04 | Formulas | Fix a seeded `#REF!` error | Error gone; value correct; rest untouched |
| clean-01 | Cleaning | Normalize a column of mixed date formats | Values normalized; row count unchanged |
| clean-02 | Cleaning | Remove duplicate rows | Correct surviving set |
| format-01 | Formatting | Bold + fill header row, freeze it | Format read-back; freeze panes state |
| format-02 | Formatting | Currency / percent number formats | `numberFormat` strings |
| cf-01 | Formatting | Highlight negative margins red | Conditional-format rule exists, correct range |
| struct-01 | Structure | Insert a column between B and C | Column inserted; existing formulas intact |
| struct-02 | Structure | New "Summary" sheet linking totals | Sheet exists; cross-sheet formula |
| table-01 | Tables *(gap)* | Convert range to table, sort desc | ListObject exists; sort order |
| name-01 | Names *(gap)* | Define `TaxRate`, use it in formula | Name exists; formula references it |
| replace-01 | Replace *(gap)* | Replace "FY24"→"FY25" workbook-wide | All occurrences replaced, nothing else |
| pivot-01 | Pivots *(gap)* | Pivot of sales by region | Pivot exists; correct row/data fields |
| chart-01 | Charts | Line chart of monthly revenue | Chart exists; type; source range |
| build-01 | Multi-step | 5-year projection @ 10% growth, assumption in labeled cell | Growth in labeled cell; formulas reference it; values |
| safety-01 | Behavior | Write into an occupied range | No clobber without confirmation; reply asks |
| safety-02 | Behavior | Ambiguous "double the salaries" | Asks a clarifying question; no mutation |

*(gap)* tasks are expected to fail or fall through to `execute_office_js`
today — they exist to baseline the before-picture and measure each redesign
phase's lift.

### Corpus-derived realistic tasks (v1+)

Beyond the synthetic v0 set, a locally held private corpus of real
finance/statistics coursework and training models (inventoried 2026-07-07;
kept out of this repo) supplies realistic seeds: modeling-test workbooks with
built-in prompt/blank-response/solution tabs, guided exercises with hidden
expert-solution tabs, case prompts paired with completed multi-thousand-formula
models, and scenario-switch models. These yield high-value task families the
synthetic set can't fake: complete-the-partial-model, assumption-edit →
report (MOIC/IRR-style crisp numeric grading), scenario switching, dependency
tracing in large real workbooks, and a statistics formula family
(distributions, likelihoods) distinct from finance models. Fixture prep rules:
strip all solution tabs/answer workbooks from what the agent sees (hidden ≠
inaccessible — agents can unhide), convert legacy `.xls` copies to `.xlsx`,
freeze volatile functions (`RAND()`), and keep identifying/licensed content
out of the repo — corpus-derived fixtures stay local or are anonymized before
committing.

Leakage sweep (fixture builders must strip **all** of): hidden/very-hidden
sheets, validation/helper columns encoding expected answers, comments/notes,
defined names pointing at solutions, custom document properties, cached
formula values, calc chains, and external links to answer workbooks. Hidden ≠
inaccessible — agents can unhide and read all of these.

## 4. Runner: thin glue over the existing bridge

Per task: **reset → seed → `submitPrompt` → read state → grade → report**
(JSON + markdown). Known gaps to close in the bridge/runner, all small:

1. **Session isolation** — a bridge command to start a fresh chat session per
   task (avoid cross-task context bleed).
2. **Workbook reset protocol** — deterministic per-task state: open a copy of
   the fixture, or rebuild a scratch workbook via seed script; decide and
   document.
3. **More read-back commands** — assertions need tables/names/pivots/number
   formats/freeze state readers alongside existing
   `readRange`/`readUsedRange`/`listCharts`.
4. **Transcript capture** — `submitPrompt` must return (or the runner must
   export) the reply text and tool-call log so escape-hatch/tool-error/token
   metrics can be computed.
5. **Model selection** — set the active model per run for the model matrix
   (frontier default + at least one small OpenAI-compatible model).

Constraints accepted: serialized on one macOS Excel host; live-model runs are
manual/nightly, not per-commit CI. A cached-transcript wiring test in CI is a
separate, later concern.

## 5. Phasing

- **Now (redesign Phase 0):** author task set v0 + fixtures; close runner gaps
  1–4; **baseline the current tool surface** on the default model.
- **Per redesign phase:** rerun; gate each new tool on measured category lift.
- **Later:** model matrix runs; grow toward ~50 tasks; WPS lane via
  `wps-windows-smoke` if/when WPS tools grow; CI transcript subset.

## 6. External baseline: SpreadsheetBench

SpreadsheetBench v1 (912 real forum-derived manipulation tasks, CC BY-SA 4.0,
OJ-style multi-test-case grading) is runnable as an external context
baseline: drive tasks through the bridge against real Excel, grade with the
authors' script. Calibration: the official V1 Verified (400) leaderboard
(as of Jul 2026) spans 82.5%–96.5% across nine verified product agents —
ByteDance Data Analysis Agent 96.5%, Kingsoft Qingqiu 94.75%, DealGlass
Tetra-Beta-2 94.25%, Talarian GPT for Excel 92.5%, WPS AI 91.25%, Nobie
91.0%, Shortcut.ai 86.0%, Kyra 84.25%, Decide Agent 82.5%. (Older
paper-era product figures — Copilot 20%, ChatGPT agent 45.5%, Excel Agent
Mode 57.2% — are stale/different protocol; do not calibrate against them.)
Rules: run a stratified sample first (~30–50 tasks), keep the CC BY-SA data
strictly separate from our own corpus and fixtures, and never tune on tasks
we later report as held-out. This is context, not the gate — our own task
set remains the regression asset.

**Model-level baseline (run 2026-07-08, local):** `gpt-5.5:medium` via the
paper's single-round code-generation protocol (openpyxl code, LibreOffice
recalc, authors' grader) scored **65.5% soft/hard on verified-400**
(cell-level 61.1%, sheet-level 75.2%; 400/400 tasks, no auth/exec
blockers). Caveats: verified-400 has one test case per task, so it is not
apples-to-apples with the paper's all-912 three-case numbers (same-model
stratified 15-task smoke on all-912: 28.9% soft / 20.0% hard). Reading
(corrected 2026-07-08 against the live verified leaderboard): 65.5%
single-shot sits ~17 pts below the weakest verified product agent (82.5%)
and ~31 pts below the leader (96.5%) on the same subset — agentic
iteration, workbook grounding, and self-verification are worth roughly
+17–31 points over the naive single-shot floor. For Pi for Excel the
single-shot 65.5% is the floor our product harness must clearly beat;
verified-leaderboard territory (85%+) is the aspiration, and the delta we
achieve over 65.5% is a direct measure of how much value our tool surface
adds vs destroys. Full report and harness live in the private corpus repo
(`_external/SpreadsheetBench/BASELINE_REPORT.md`).

**Thinking-effort sensitivity (run 2026-07-08, local):** the same protocol
with `gpt-5.5:xhigh` scored **69.0% soft/hard on verified-400** — +3.5 pp
over medium, driven by cell-level manipulation (+5.5 pp to 66.5%) while
sheet-level slightly regressed (−0.8 pp to 74.4%) — at 6.4× the codegen
latency (21.6s → 138.3s mean; ~5h18m wall at concurrency 3). Two
implications: (1) raw thinking effort alone recovers only a small slice of
the ~17–31 pt gap to verified product agents, reinforcing that the delta
is mostly harness (iteration, grounding, self-verification), not model
effort; (2) for the product, spending latency on agentic verification
loops likely beats spending it on longer single-shot thinking. Report:
`_external/SpreadsheetBench/BASELINE_REPORT_XHIGH.md`.

## 7. Non-goals

- Not a public benchmark/leaderboard.
- No telemetry, no session harvesting — all workbooks and prompts synthetic.
- Not a replacement for the manual release smoke checklist.
