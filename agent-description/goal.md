# Goal: complete the Pi for Excel agent description

You are working in `agent-description/` of the pi-for-excel repo. Read `README.md`, `glossary.md`, `foundations/the-tool-contract.md`, and `tools/reading-a-range.md` first. The README defines the purpose, the variant rows, the hostility list, the cross-cutting order, the method, the structure and the coverage table. The other three are the exemplars: match their depth, tone and structure. Your job is to keep every document in the README's structure accurate, extend it when tools change, and keep [bug-triage.md](bug-triage.md) ranked.

## Source of truth

The pi-for-excel source is the rest of this repository (everything outside `agent-description/` and `product-description/`). Describe the agent of one Pi tab in the default configuration: Auto execution mode, no extensions, no experimental features, no rules or conventions, one provider. WPS, extension tools and the local bridges are out of scope (see the README's scope decisions).

For each document, read in this order before writing:

1. The tool: `src/tools/{tool}.ts` (schema descriptions, result text, error paths), `src/tools/capabilities.ts` (its prompt line), `src/tools/execution-policy.ts` (read or mutate, impact).
2. The wrappers every call passes through, outermost first: `src/tools/output-truncation.ts`, `src/tools/with-connection-preflight.ts`, `src/tools/with-workbook-coordinator.ts`, `src/workbook/coordinator.ts`; and the agent loop in `node_modules/@earendil-works/pi-agent-core/dist/agent-loop.js` (`prepareToolCall`, `executeToolCalls`, `transformContext`).
3. What the agent is told around the call: `src/prompt/system-prompt.ts`, `src/taskpane/context-injection.ts`, `src/context/`, `src/messages/convert-to-llm.ts`, `src/messages/tool-result-shaping.ts`.
4. Backups: `src/tools/mutation/`, `src/tools/recovery-metadata.ts`, `src/workbook/recovery-log.ts`.
5. Tests: `tests/workbook-write-inspect-undo-contract.test.ts`, `tests/tools-output-truncation.test.ts`, `tests/context-injection.test.ts`, `tests/execution-policy.test.ts`, `tests/with-workbook-coordinator.test.ts`, and the per-tool tests. There are no tests for history shaping, the selection context or the change list.
6. Then probe the real tool in Excel on the web (see [verification/README.md](verification/README.md)). A document is not finished until its hostile states have been tried.

## Document skeleton

Every tool document uses all eight sections, in this order, with these headings. Foundations and hostility documents may drop sections that do not apply but must still cover the hostility list wherever a tool call is involved. A tool document runs 120–250 lines.

The H1 is the tool's name in code for tool documents ("`read_range`", "`write_cells`" or "`format_cells` and `conditional_format`"), a noun phrase for foundations ("The tool contract") and a descriptive phrase for hostility ("Cell-edit mode", "Payload limits and large ranges").

1. `## Summary` — what the tool lets the agent do; its prompt line quoted; read or mutating; whether it is in Confirm-mode approvals.
2. `## The simple case` — the common call and the exact result text, in a fenced block.
3. `## The call, event by event` — a Mermaid `stateDiagram-v2`, then five `###` subsections in this order: **Forms the call**, **Rejected at once**, **Touches Excel**, **While running**, **The result**.
4. `## Workbook and host state` — a table `State | True before the call | Changes while running` with exactly these rows: Sheet protection; Calculation mode; Merged cells; Tables, filters and hidden rows; Data validation; Dynamic arrays and spills; External links; Very large used range; Number formats and locale; Host and API set; Execution mode. Every cell filled; "No effect." where that is the answer.
5. `## Hostility and interrupts` — a table `Event | Before Excel is touched | While running` with exactly these rows: The user is editing a cell; The user changes the selection or the active sheet; The user changes data the call depends on; Another tab's write holds the workbook lock; Excel throttles, times out or rejects the payload; Recalculation is pending; The workbook is co-authored; Autosave or a save boundary; The user presses Stop; The context fills. Every cell filled. Then a sentence on what the agent is left knowing.
6. `## Interactions with other systems` — one bold-led paragraph per concern, in this order, even when the answer is "No interaction.": **What the agent is told.** **Verification.** **Backups and recovery.** **The workbook lock and parallel calls.** **Execution mode and approvals.** **Truncation and history shaping.** **Auto-context.** **Prompt caching.** **Host differences.** **What the user sees.**
7. `## Edge cases` — a bullet list.
8. `## Open questions and verification` — bullets: what was confirmed and how; what is read from code only; what looks like a bug (link the triage entry); then ``Verified against pi-for-excel commit `{sha}`.``

## Writing rules

- Describe the agent's experience, not the code. Quote result text exactly; it is the experience.
- Say how each claim is known: "confirmed" (observed in Excel on the web) or "read from code".
- Technical detail only in `> Technical note:` block quotes, and only when the mechanism changes what the agent would expect.
- Use the glossary's words. Add a term to `glossary.md` with a full definition before using it. Never coin a synonym.
- Sentence case for headings. Direct, concrete language.
- Surprising behavior is stated plainly, with the reason if the code or a comment gives one. If it looks like a defect, say so and add or link a `B-NN` entry in [bug-triage.md](bug-triage.md).
- Cross-reference instead of repeating. The foundations own the shared facts: [the tool contract](foundations/the-tool-contract.md) owns result shapes; [auto-context](foundations/auto-context.md) owns the injected note; [limits](foundations/limits.md) owns every number; [the workbook as the agent sees it](foundations/the-workbook-as-the-agent-sees-it.md) owns addresses and values; each [hostility](hostility/) document owns its family. The user's view belongs to [product-description/](../product-description/README.md).

## Things already established (do not re-derive, do not contradict)

- Source commit for this pass: `ec453df` (the same as product-description). `origin/main` has since gained `e74e2ea`, which changes `fill_formula` on the web (2D targets filled row then column; single-cell targets no longer autofill onto themselves); documents say where that matters.
- Every workbook tool reports Excel failures as error-shaped results (text "Error …", `isError` false). Only schema validation, a declined approval, an abort before start and unsupported-host calls are thrown errors.
- Details never reach the model. Backup IDs, change samples and truncation metadata are details; the backup note and the "Changed cell(s): N." line are text.
- Read calls never wait for the workbook lock, so a read in the same batch as a write can see the old value (confirmed).
- Tool calls in one batch run in parallel (the agent loop default); the OpenAI-style request sets `parallel_tool_calls: true`.
- Auto-context is built before every model call, placed just before the latest real user message, and never stored. The blueprint is re-sent on every call labelled "[Workbook context refresh: context missing from history]" after the first. The change list appears only in the first model call after a change (confirmed).
- The change list reports Pi's own writes as user changes (Excel reports them with `triggerSource` "ThisLocalAddin"; user edits are "Unknown"), never tracks sheets added after the taskpane opened, and is one list for all tabs.
- The selection context has no size cap: a 5,000-row selection on a 5-column sheet produced a 311,761-character note on every model call (context meter 3% → 32% of 500k for "Reply with just OK."). A whole-column or whole-sheet selection produces no selection context at all.
- Values are raw: dates are serials, percentages fractions, currency plain numbers. Zero renders as `0.0000e+0`. Numbers below 1e-6 in magnitude render in exponent form.
- Payload limit on Excel on the web: `read_range` of 10,000 × 5 succeeded; 50,000 × 5 failed; any whole-column read (`A:A`) failed even on a 4-row sheet; `search_workbook` fails for the whole workbook once any sheet's used range is too big; `get_workbook_overview` with `sheet` fails on a big sheet.
- False failures confirmed: `fill_formula` over 200,000 rows and `write_cells` of 60,000 × 4 both reported the payload error after the change was applied, with no backup.
- Silent successes confirmed: writes into merged areas ("Written to **Merged!B1**" with nothing written; 2×2 into a merged block landed 1 of 4); overwriting a spill child (result clean, anchor now `#SPILL!`); invalid values under data validation accepted; formulas under manual calculation "verified" as 0.
- Cell-edit mode on the web fails every batch at once (3–12 ms); the agent stopped and asked the user to press Enter (confirmed in a real turn).
- Protected sheets: "You cannot perform the requested operation." for `write_cells`, `format_cells`, `modify_structure` and `execute_office_js`; no mention of protection. Unlocked cells on a protected sheet are writable.
- Unqualified addresses go to the active sheet at the moment of the batch (confirmed).
- Concurrent Confirm-mode approvals: opening a second approval resolves the first as declined without a click (confirmed at the dialog level).
- Truncation: 50 KB / 2,000 lines, head for workbook tools, tail for `python_run`, `tmux`, `mcp`, `execute_office_js`; `execute_office_js` additionally cuts its JSON at 8,000 characters. History shaping keeps the 6 most recent tool results whole (fewer below a 128k window, floor 2).
- Backups are skipped above 20,000 cells; the note then says "Backup capture was skipped (workbook identity unavailable or snapshot data unsupported)" in details and "ℹ️ Backup not created for this mutation. …" in text, whatever the real reason.
- Model used for real turns: GPT-6.1 Sol via `openai-codex`, 500k window, thinking High.

## Working rules

- Commit with `docs: add agent-description/{path}` or `docs: revise agent-description/{path}`. This repo does not add AI attribution to commit messages.
- Do not modify anything outside `agent-description/`.
- Do not add files outside the README's structure without updating the structure and coverage table.
- When a behavior cannot be determined, write what you could determine, put the rest in "Open questions", and move on.
- Never renumber checklist IDs or triage IDs once used.
