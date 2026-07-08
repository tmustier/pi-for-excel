# Learnings from the first live eval cycle (2026-07-07/08)

Sources: three live product runs (`wso-basic-lbo-01` on opus-4-8;
`gpu-farm-doctor-01` ×2 on gpt-5.5/high) plus the SpreadsheetBench v1
model-level baseline (`gpt-5.5:medium`, verified-400, 65.5% soft/hard).
Companion docs: [`../proposals/agent-evals.md`](../proposals/agent-evals.md),
[`../proposals/agent-tool-interface-redesign.md`](../proposals/agent-tool-interface-redesign.md).

## A. Product defects surfaced (agent behavior)

### A1. No edit-restraint norm → collateral edits (P0)
`gpu-farm-doctor-01`: gpt-5.5/high fixed all 3 injected bugs exactly, then
made **31 unintended cell edits** — rewired `Calculations!I167:M171` refs
from healthy rows 118:122 to 82:86 (corrupting 2029–30 outputs), widened a
subtotal (`Statements!I51`), and restyled tie-checks to `ABS()` form.
Score: Modif. 3/3, Acc. 27/35 — the exact `Modif. high / Acc. low`
over-editing signature SpreadsheetBench documents.

The system prompt WORKFLOW has "read first", "verify writes", "overwrite
protection" — but **no scope norm**. Nothing says: make the smallest set of
edits that satisfies the request; report suspicious cells instead of
silently fixing them; don't restyle working formulas. Overwrite protection
is no defense here — a repair task legitimately authorizes overwrites, and
once in "fix mode" the agent applies it broadly.

**Fix:** add explicit edit-scope norms to WORKFLOW (small static addition,
cache-safe). Then re-run the doctor task to measure the delta — first real
A/B for the eval loop.

### A2. No change-report norm → scope creep is invisible (P0)
The agent never enumerated what it changed, so neither the user nor the
agent itself caught the 31-cell creep. Tool-level mutation receipts exist;
what's missing is the norm that the agent's final answer aggregates them
("changed: K42, I156, L97 — nothing else"). For finance users this is
also the audit-trail feature that builds trust.

**Fix:** WORKFLOW norm now; later, consider a session-level change ledger
(receipt aggregation) the UI can render — ties into the redesign proposal's
mutation receipt contract (§5.2).

### A3. Write verification is mechanical, not semantic (P0/P1)
"Verify writes" today = the write landed without host errors. It does not
mean "the model output moved the way the fix implies". Evidence from both
lanes: doctor collateral damage went unnoticed (downstream 2029–30 outputs
silently inflated ~19–400%); SpreadsheetBench error taxonomy shows
**127/400 failures were 'code ran, output existed, state wrong'** — the
single-shot analog of the same gap.

**Fix:** norm: after formula repairs/structural edits, re-read the affected
downstream outputs and sanity-check direction/magnitude before declaring
done. `trace_dependencies` already exists — the capability is there, the
prompt never tells the agent to close the loop with it.

### A4. Reference rewiring without label verification (P1)
Root-cause hypothesis for the I167:M171 rewire: rows 82:86 and 118:122 are
two similarly-labeled blocks; the agent pattern-matched the wrong one and
"corrected" healthy references. A norm ("before rewiring a reference,
verify the row/column label of both old and new targets") plus semantic
addressing (redesign §3.3) both attack this.

## B. Strategic learnings

### B1. Iteration is the product's value, and we can now measure it
Single-shot gpt-5.5:medium code-gen: **65.5%** on verified-400. Verified
product agents: **82.5–96.5%**. The +17–31pt band *is* agentic iteration,
workbook grounding, and self-verification. Consequence: the highest-value
next eval is a **stratified SpreadsheetBench sample through the real
product** (taskpane + real Excel) to place Pi on that ladder. If Pi lands
near the single-shot floor, the harness is destroying model value; the gap
decomposition tells us exactly where to work.

### B2. `unintended_edited_cells` should be a first-class metric everywhere
The run-4 grading built a seed-vs-final formula diff (openpyxl seed × bridge
snapshot) that cleanly separates target fixes from collateral edits. Adopt
in every lane, not just doctor tasks — destructive edits are the top trust
risk for real users.

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
6. **pi CLI as a pure model endpoint** needs the full flag set
   (`-p --no-tools --no-context-files --no-extensions --no-skills
   --no-prompt-templates --no-themes --no-session`) — bare `pi -p` is a
   full agent with side effects.

## D. Priority queue

| # | Item | Type | Cost | Evidence |
|---|---|---|---|---|
| 1 | Edit-scope + change-report + semantic-verify norms in system prompt | product | S | A1–A3 |
| 2 | Re-run `gpu-farm-doctor-01` post-change; measure `unintended_edited_cells` delta | eval | S | A1 |
| 3 | `unintended_edited_cells` standard in grading | eval | S | B2 |
| 4 | Bridge: `newSession` + `setModel` + usage export + `status` enrichment | infra | M | C1–C2 |
| 5 | SpreadsheetBench stratified sample through the real product | eval | M | B1 |
| 6 | Semantic addressing + `trace_dependencies` workflow integration | product | L | A4, redesign §3.3 |
