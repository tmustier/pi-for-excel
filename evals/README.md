# Agent evals harness

Thin tooling for running and grading **real-Excel** evals of the pi-for-excel
agent, per [docs/proposals/agent-evals.md](../docs/proposals/agent-evals.md).
The task set is the asset; this harness stays deliberately thin.

## Repo / corpus split

| Lives here (public repo) | Lives in the private corpus (local-only) |
|---|---|
| Grader, bridge helpers, shared libs | Seed workbooks + fixture builders |
| Task **schema** + example spec | Real task YAMLs (some derived from licensed training materials) |
| Proposal + research docs | Run reports, snapshots, expected-value oracles |

The private corpus is a local-only git repo at `~/projects/excel-eval-corpus`
(never push; licensed/personal source material is excluded by a whitelist
`.gitignore` there as defense-in-depth).

## Components

- **`bin/bridge.sh`** — token-efficient CLI over the background-verification
  bridge (`scripts/`-served taskpane + tokened loopback HTTPS server). Lists
  live taskpane clients, polls one-line run status (~200 bytes vs ~100 KB for
  a visual observe), and issues raw bridge commands (`readRange` with
  `include:"all"`, `submitPrompt`, ...). This is the canonical way to monitor
  a live eval run; visual observation is for one-off disputes only.
- **`bin/grade.py`** — standard grader. Inputs: seed xlsx, per-sheet bridge
  snapshot JSONs, expected-values JSON, optional target-fix map, protected
  sheets. Outputs a JSON verdict + human summary with four checks:
  1. `cells_match` — graded output cells vs oracle (`--rel-tol` for nonzero
     expecteds, `--abs-tol` at zero; bools must be bools; addresses outside
     the snapshot used range are hard failures, never silently indexed)
  2. `target_fixes` — quote-aware normalized formula equality for intended
     edits (case/whitespace normalized only outside string/sheet literals)
  3. `no_mutation` — protected sheets untouched: values AND formulas,
     diffed over the union of seed and snapshot cells (whitespace-only
     spacer strings ≡ empty, documented leniency)
  4. `unintended_edited_cells` — formula-level diffs vs seed outside the
     intended edit set, incl. formula→value replacement and formulas
     deleted/cleared even when the final used range shrank. First-class destructive-edit metric: an agent can fix all
     target bugs (`Modif.` pass) while silently rewiring healthy formulas
     (`Acc.` fail). Observed in practice: 31 unintended edits in a run that
     fixed 3/3 target bugs.
- **`lib/lo_recalc.py`** — LibreOffice headless recalc for oracle generation
  (openpyxl round-trips drop cached values → LO computes fresh). Uses an
  isolated LO user profile per call so concurrent soffice instances are
  undisturbed. **Fidelity contract:** validated 35/35 (rel tol 1e-6) against
  an Excel-derived oracle on a model using CHOOSE/IF/ISERROR/IFERROR/SUM/
  SUMIF/SUMPRODUCT/MIN/MAX/AVERAGE; re-validate before trusting it for
  dates, IRR/NPV, lookups, or text functions.
- **`lib/scrub.py`** — leakage scrub + zip-level `assert_no_leakage()`
  gate that fixture builders run on every emitted seed: hidden sheets,
  personal metadata, custom props, comments, external links, calc chain,
  cached formula values, VBA, and non-builtin defined names. (Real catch:
  it flagged training-workbook external links pointing at a
  `*_CorrectAnswers_*.xls` path, tutoring comments, and ~130 defined names
  of validation machinery.)

## Task spec schema

Tasks are data (YAML), graded on **workbook state**, not transcript claims.
See `tasks/example-doctor.yaml`. Key doctrines:

- **Two-level scoring** (SpreadsheetBench pattern): `Modif.` = intended
  changes made; `Acc.` = full workbook correctness. Report both.
- **Hidden variants**: builders emit variants (a/b/c) with perturbed input
  literals so memorized/hardcoded answers fail across variants. Report the
  variant used per run; track variant pass rate.
- **Placement graded semantically** unless the prompt pins exact cells.
- **Budgets**: max tool calls / minutes per task.

## Running a live eval (current manual protocol)

1. Serve the taskpane (`npm run dev`, port 3141) and start the
   background-verification bridge (see
   `.agents/skills/excel-background-verification/`).
2. Copy the seed workbook to a scratch path; open in Excel (`open -g`).
3. Open the Pi taskpane in that workbook window (semantic AX press only).
4. `bridge.sh clients` → target the new client id; set model/session.
5. Submit the task prompt via bridge `submitPrompt`.
6. `bridge.sh watch <clientId>` until idle.
7. Snapshot sheets: `bridge.sh cmd <id> readUsedRange '{"sheet":"...","include":"all"}'`
   → save one JSON per sheet.
8. `bin/grade.py --seed ... --snapshots ... --expected ... [--targets ...]
   [--no-mutate Sheet]` → verdict JSON + summary.

Known gaps to automate next (tracked in the proposal): per-task fresh chat
session, workbook reset protocol, transcript/usage export, formula-level
read in one call (readUsedRange include:"all" covers it), run manifest.

## External calibration

SpreadsheetBench v1 verified-400, single-round code-generation protocol
(model-level floor, not product numbers): `gpt-5.5:medium` 65.5%,
`gpt-5.5:xhigh` 69.0% (+3.5 pp at 6.4× codegen latency). Live verified
leaderboard product agents span 82.5–96.5%. The product harness must
clearly beat the single-shot floor; the delta is a direct measure of what
the tool surface adds. Full reports live in the private corpus.
