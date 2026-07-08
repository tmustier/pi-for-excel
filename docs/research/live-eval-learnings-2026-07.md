# Learnings from the first live eval cycle (2026-07-07/08)

Sources: live product runs (`wso-basic-lbo-01` on opus-4-8;
`gpu-farm-doctor-01` on gpt-5.5/high across hidden variants), the corrected
local gpu-farm fixture/oracle loop, and SpreadsheetBench v1 model-level
baselines (`gpt-5.5:medium` 65.5%, `gpt-5.5:xhigh` 69.0% on verified-400).
Companion docs: [`../proposals/agent-evals.md`](../proposals/agent-evals.md),
[`../proposals/agent-tool-interface-redesign.md`](../proposals/agent-tool-interface-redesign.md).

## A. Product defects surfaced (agent behavior)

### A1. First-order lesson: clean-source audit before hard over-edit metrics (P0)
The first gpu-farm doctor runs looked like a classic `Modif. high / Acc. low`
over-editing failure: the model fixed the 3 planted bugs, then changed
`Calculations!I167:M171`, `Statements!I51`, and `Statements!I71:M71`, which
the initial oracle counted as unintended.

A source audit overturned that interpretation. Those edits were legitimate
source-quality repairs in the supposedly clean workbook:

- `Calculations!I167:M171` is the server utilization cohort and should use
  the server useful-life schedule (`82:86`), not the facility useful-life
  schedule (`118:122`).
- `Statements!I51` should include dividends consistently (`SUM(I47:I50)`).
- balance-check row 71 should use `ABS(total L&E - total assets) < 1`.

After normalizing those source issues before injecting the 3 doctor-task
bugs, the supposedly failing run4 snapshots regrade as a strict PASS:
**35/35 cells, 3/3 target fixes, Inputs unchanged, 0 unintended edits**.
Current `main` also passes hidden variant b under the normalized fixture.

**Fix shipped in corpus:** bake unrelated source-quality repairs into the
fixture baseline and keep `unintended_edited_cells` strict only after the
clean seed itself is audited.

### A2. Change reports are still the user-facing audit layer (P0)
The early oracle was wrong, but the UX lesson remains: after repair work,
the assistant must enumerate exactly what it changed and why. Without that,
valid source-quality repairs, suspicious-but-unfixed cells, and accidental
collateral edits all look the same to a finance user.

Tool-level mutation receipts exist; what is missing is the norm that the
final answer aggregates them (for example: "changed: K42, I156, L97; also
normalized I167:M171 after verifying server useful-life labels"). For finance
users this is the audit-trail feature that builds trust.

**Fix:** WORKFLOW norm now; later, consider a session-level change ledger
(receipt aggregation) the UI can render — ties into the redesign proposal's
mutation receipt contract (§5.2).

### A3. Write verification is mechanical, not semantic (P0/P1)
"Verify writes" today = the write landed without host errors. It does not
mean "the model output moved the way the fix implies". The gpu-farm loop
showed both sides of the issue: a formula diff looked dangerous until a
semantic source audit proved it was correct; SpreadsheetBench's error
taxonomy shows the converse at scale — **127/400 failures were 'code ran,
output existed, state wrong'**.

**Fix:** norm: after formula repairs/structural edits, re-read the affected
downstream outputs and sanity-check direction/magnitude before declaring
done. `trace_dependencies` already exists — the capability is there, the
prompt should make the agent close the loop with it.

### A4. Reference rewiring can be valid even when it looks like collateral damage (P1)
The `I167:M171` case is the cautionary example. It looked like broad
reference rewiring, but label/dependency inspection showed it was a valid
fix. Future grading and product UX should distinguish:

1. **auditable reasoning:** did the agent inspect old/new target labels and
   downstream impact before rewiring?
2. **mutation safety:** did it preserve the normalized seed outside genuine
   fixes?

Semantic addressing and dependency-aware mutation receipts (redesign §3.3 / §5.2)
remain useful, but the eval lesson is: inspect the workbook semantics before
assuming a formula diff is bad.

## B. Strategic learnings

### B1. Iteration is the product's value, and we can now measure it
Single-shot gpt-5.5:medium code-gen: **65.5%** on verified-400. Verified
product agents: **82.5–96.5%**. The +17–31pt band *is* agentic iteration,
workbook grounding, and self-verification. Consequence: the highest-value
next eval is a **stratified SpreadsheetBench sample through the real
product** (taskpane + real Excel) to place Pi on that ladder. If Pi lands
near the single-shot floor, the harness is destroying model value; the gap
decomposition tells us exactly where to work.

### B2. `unintended_edited_cells` is valuable, but only with an audited clean seed
The grader's seed-vs-final formula diff is still the right trust metric for
repair tasks. But gpu-farm showed the failure mode: if the seed contains
unrelated plausible bugs, a strict formula diff punishes good work. Adopt the
metric everywhere *with* fixture hygiene: hidden tabs/answers stripped,
cached values removed, and clean-source formulas audited for obvious
inconsistencies before defining the target diff.

## C. Eval-infra fixes (dev-side, from running the harness)

1. **Bridge additions wanted:** `newSession` (session hygiene per run),
   `setModel`/`setThinkingLevel` (currently AX hacks), transcript/usage
   export (token + tool-call metrics per run), and a long-horizon
   `waitUntilIdle` (the current 120s clamp + sawProgress race is unusable —
   poll `status` instead).
2. **Cheap observation is solved and should be codified:** bridge `status`
   is ~200B vs ~100k chars for a taskpane `observe` screenshot. Helper:
   corpus `_evals/bin/bridge.sh` (clients/status/watch/cmd). Adding
   `lastToolCall` + last-assistant-snippet to `status` would remove the
   remaining need for visual observation.
3. **Runner preflight is mandatory:** dead Codex OAuth + stopped proxy
   produce silent instant-fail runs (msgs=2). Cheap canary prompt before
   counting any run.
4. **HMR isolation:** the eval Vite server must run from a worktree no
   other agent edits; attempt 1 of the doctor run was killed by unrelated
   dev edits triggering taskpane reloads.
5. **Client IDs churn on every reload** — resolve via `/health` lastSeenAt +
   `status` workbookName; never cache across reloads.
6. **Vite must be started with background-verify env** — if
   `VITE_PI_BACKGROUND_VERIFY_URL/TOKEN` are missing, the taskpane looks alive
   but never registers with the bridge.
7. **pi CLI as a pure model endpoint** needs the full flag set
   (`-p --no-tools --no-context-files --no-extensions --no-skills
   --no-prompt-templates --no-themes --no-session`) — bare `pi -p` is a
   full agent with side effects.

## D. Priority queue

| # | Item | Type | Cost | Evidence |
|---|---|---|---|---|
| 1 | Keep edit-scope + change-report + semantic-verify norms (#635) | product | done | A2–A3 |
| 2 | Clean-source audit/normalization before strict formula-diff grading | eval | done for gpu-farm | A1, B2 |
| 3 | Expand hidden-variant pass rate beyond regraded a + live b/c evidence with fresh normalized runs if needed | eval | S | A1 |
| 4 | `unintended_edited_cells` standard in grading, with fixture hygiene gate | eval | S | B2 |
| 5 | Bridge: `newSession` + `setModel` + usage export + `status` enrichment | infra | M | C1–C2 |
| 6 | SpreadsheetBench stratified sample through the real product | eval | M | B1 |
| 7 | Semantic addressing + `trace_dependencies` workflow integration | product | L | A4, redesign §3.3 |
