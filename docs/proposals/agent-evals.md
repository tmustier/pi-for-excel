# Proposal: Agent Eval Suite

**Status:** Proposal (not yet accepted) — v2, restructured task-set-first
**Date:** 2026-07-07
**Companion to:** [`agent-tool-interface-redesign.md`](./agent-tool-interface-redesign.md)

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
   workbook lane. Rejected: it is a simulator-building project that would
   consume the effort budget, cannot compute formulas without bolting on an
   engine (HyperFormula is GPL; MIT parsers are partial), and produces false
   confidence exactly where fidelity matters. Unit tests keep using small
   ad-hoc fakes; evals use real Excel. Revisit only if we someday need
   CI-scale or cross-platform runs.
4. **Deterministic grading.** Assert final workbook state: cell values,
   formula text, number formats, object existence (tables/charts/pivots/
   names), freeze panes, and absence of leftover artifacts. Reply-text checks
   are `contains`-style and coarse. LLM-judge only ever as a labeled
   spot-check, never the gate.

## 2. Metrics (per task × model)

| Metric | Why |
|---|---|
| **Task success** (assertions pass) | The material outcome |
| **Escape-hatch rate** (`execute_office_js` calls) | Which structured tool to build next — replaces telemetry |
| **Tool-error rate** (failed calls, schema rejections, retries) | Schema/semantics friction |
| **Efficiency** (tool calls, tokens, wall time) | Batch/high-level tool gaps |
| **Behavioral checks** (asked-when-ambiguous, no-clobber-without-confirm) | Interaction quality |

Per-tool-surface phase, the acceptance criterion for a new tool is a measured
drop in escape-hatch/failure rate in its category.

## 3. Task set v0 (~20 tasks)

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

## 6. Non-goals

- Not a public benchmark/leaderboard.
- No telemetry, no session harvesting — all workbooks and prompts synthetic.
- Not a replacement for the manual release smoke checklist.
