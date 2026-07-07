# Proposal: Agent Eval Harness

**Status:** Proposal (not yet accepted)
**Date:** 2026-07-07
**Companion to:** [`agent-tool-interface-redesign.md`](./agent-tool-interface-redesign.md)

Telemetry is off the table (product decision), and no historical usage corpus
exists. Evals are therefore the evidence loop for the tool-surface redesign —
and, once built, a regression gate for prompt/context/tool changes and the only
honest way to answer "does this work on smaller models?" (#603 gateway users).

## 1. What we measure

Per task, per model:

| Metric | Why |
|---|---|
| **Task success** (deterministic assertions on final workbook state) | The material outcome |
| **Escape-hatch rate** (`execute_office_js` calls per task) | Which structured tools are missing — replaces telemetry |
| **Tool-error rate** (failed/retried calls, schema rejections) | Schema and semantics friction |
| **Efficiency** (tool calls, tokens, wall time) | Batch/high-level tool gaps |
| **Behavioral checks** (asked-when-ambiguous, proceeded-when-clear, overwrite-protection flow followed) | Interaction quality |

Grading is deterministic wherever possible: assert cell values, formula text,
number formats, object existence (charts/tables/pivots/names), and absence of
leftover artifacts. LLM-judge only for soft criteria (reply quality), and only
as a spot-check, never the primary gate.

## 2. Architecture: two lanes

### Lane A — fast, deterministic (CI-able)

Node harness running the **real agent loop** (pi-ai) with the **real tool
implementations** against an **in-memory fake workbook**:

- Grow the per-test ad-hoc fakes (e.g. `withFakeExcel` in
  `tests/charts-tool.test.ts`) into a shared `tests/helpers/fake-workbook.ts`
  that models sheets, ranges, values, formulas, formats, and object collections.
- Formula evaluation is the hard part. Options, in preference order:
  1. Assert on **formula text + inputs** rather than computed values where
     possible (most tasks).
  2. Embed an MIT-licensed formula engine (e.g. `fast-formula-parser`) for
     tasks that need computed values, accepting fidelity gaps (HyperFormula is
     GPL — unusable).
- Runs headless with any provider key; this is where the model matrix runs.

The #605 `WorkbookAPI` layer is the clean long-term seam (swap an in-memory
backend under the same typed API). Until it exists, the global-`Excel` fake
pattern the tests already use is sufficient.

### Lane B — ground truth (real Excel)

Reuse the `excel-background-verification` harness to run a subset against real
Excel desktop on macOS: host-fidelity behaviors that fakes cannot capture
(freeze panes semantics, pivot API quirks, formatting inheritance on insert,
CSP-sensitive paths). Run pre-release and when Lane A and reality are suspected
to diverge. WPS lane later via `wps-windows-smoke` if/when WPS tools grow.

## 3. Task suite

Seed workbooks are fixtures (small `.xlsx` files or builder scripts). Task spec
is data, not code:

```yaml
id: tables-01-create-and-sort
category: tables
seed: fixtures/sales-raw.xlsx
prompt: "Turn the data on Sheet1 into a table and sort it by Revenue, highest first."
assertions:
  - table_exists: { sheet: Sheet1, min_rows: 20 }
  - sorted_by: { column: Revenue, order: desc }
budget: { max_tool_calls: 8 }
```

Categories (mirroring the target tool surface + competitor capability list):

1. **Orientation** — "what's in this workbook?", find-by-label questions
2. **Formulas** — write, fill, explain, trace; formulas-not-values discipline
3. **Data cleaning** — normalize dates, split names, dedupe (CfE's strongest suit)
4. **Formatting** — conventions compliance, number formats, conditional rules
5. **Structure** — insert/delete/move rows/cols/sheets, freeze, move/copy ranges
6. **Tables** — create, sort, filter, total row *(expected to fail today → escape hatch; measures Phase 1 lift)*
7. **Pivots / charts** — create, modify *(pivots expected to fail today)*
8. **Names / validation / replace** — gap categories, same purpose
9. **Multi-step builds** — small model build (mini-DCF), scenario run
10. **Safety behavior** — overwrite-protection flow, ambiguous-request handling,
    error recovery after a failed call

Start with ~15 tasks across categories 1–5 + a few deliberate-gap tasks (6–8);
grow toward ~50.

## 4. Phasing

- **MVP (redesign Phase 0):** harness + ~15 tasks, Lane A only, manual run,
  JSON + markdown report. **Baseline the current tool surface** — this baseline
  is the redesign's before-picture.
- **Per redesign phase:** rerun; the acceptance criterion for each new tool is a
  measured drop in escape-hatch/failure rate in its category, not vibes.
- **Later:** small deterministic subset in CI (fake model or cached transcripts
  for wiring regressions); nightly/pre-release full run with live models;
  Lane B subset in the release smoke flow.

## 5. Non-goals

- Not a leaderboard; no public benchmark claims.
- No telemetry, no session harvesting — eval workbooks and prompts are
  synthetic.
- Not a substitute for the manual release smoke test (it complements it).
