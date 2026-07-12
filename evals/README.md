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
  a visual observe), and drives a whole eval run without raw GUI input:
  - `session <clientId>` — fresh chat/session per task (`newSession`), proving
    runtime/session id + message-count reset.
  - `model <clientId> <provider> <modelId> [thinkingLevel]` — model + thinking
    switch via the production model-switch seam, validated against the model
    registry (e.g. `openai-codex gpt-5.6-sol medium`); unknown model or
    unsupported thinking level fails closed.
  - `submit <clientId> <text> [timeoutMs]` — submit prompt and wait to idle.
  - `wait <clientId> [baselineMsgs] [timeoutMs]` — pollable durable
    wait-until-idle; pass the `baselineMessageCount` from the submit response
    for exact start detection (never reports idle before the run starts, even
    on a silent instant-fail).
  - `transcript <clientId> [maxReplyChars]` — bounded transcript + usage export
    (reply text, per-tool call/error counts, token totals). By design it
    excludes raw tool-call arguments and raw user/tool-result text; it exports
    only roles, lengths, tool names/errors, assistant bounded text, stop
    reasons, and usage so terminal artifacts minimize workbook/secret leakage.
  - `cmd <clientId> <type> <payloadJson> [timeoutMs]` — raw command escape
    hatch (`readUsedRange` with `include:"all"`, `officeProbe`, ...).
  This is the canonical way to monitor/drive a live eval run; visual
  observation is for one-off disputes only.
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
4. `bridge.sh clients` → target the new client id.
5. `bridge.sh session <id>` → fresh chat/session for the task.
6. `bridge.sh model <id> openai-codex gpt-5.6-sol medium` → pin model + thinking.
7. `bridge.sh submit <id> "<task prompt>" 180000` → submits and waits to idle;
   note the `baselineMsgs` it prints. (Re-poll with
   `bridge.sh wait <id> <baselineMsgs> 180000` if you split submit/wait.)
8. `bridge.sh transcript <id>` → reply + tool-call/error counts + token totals.
9. Snapshot sheets: `bridge.sh cmd <id> readUsedRange '{"sheet":"...","include":"all"}'`
   → save one JSON per sheet.
10. `bin/grade.py --seed ... --snapshots ... --expected ... [--targets ...]
   [--no-mutate Sheet]` → verdict JSON + summary.

Now automated by the bridge (previously manual gaps): per-task fresh chat
session (`session`), registry-validated model/thinking selection (`model`),
durable wait-until-idle (`wait`/`submit`), and bounded transcript/usage export
(`transcript`). Remaining gaps to automate next (tracked in the proposal):
workbook reset protocol and a run manifest.

## External calibration

SpreadsheetBench v1 verified-400, single-round code-generation protocol
(model-level floor, not product numbers): `gpt-5.5:medium` 65.5%,
`gpt-5.5:xhigh` 69.0% (+3.5 pp at 6.4× codegen latency). Live verified
leaderboard product agents span 82.5–96.5%. The product harness must
clearly beat the single-shot floor; the delta is a direct measure of what
the tool surface adds. Full reports live in the private corpus.
