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

- Follow merged main, including #786–#799 and subsequent merged product decisions. The [product goal](../product-description/goal.md#merged-fixes) records the original fixes' provenance. Unmerged proposals are not current behavior.
- Workbook failures have readable content and reliable `isError` outcomes; details are UI metadata, not model context. See [the tool contract](foundations/the-tool-contract.md).
- Calls in a batch can run in parallel, but same-runtime reads wait for earlier registered mutations. The lock serializes writes, not co-authors or all reads in other runtimes.
- Structure, Files and other people's changes are durable conversation notes when changed. Selection is fresh per request and bounded, including whole-column selections. Own writes are excluded from the other-people change note; newly added sheets are tracked. See [auto-context](foundations/auto-context.md).
- Compact and detailed reads use displayed values, with stored-value/formula and workbook-state annotations. Large reads page and bound each response; large search and overview operations avoid one unbounded values load.
- Checked writes refuse non-anchor merged targets, spill obstacles and protected targets; large writes chunk and report verification limits. A partial failure does not imply rollback. Manual-calculation read-back is not presented as proof of recalculated correctness.
- Workbook tools wait for cell-edit mode for up to 60 seconds; context collection does not block a model request on that wait. Reads may resolve an unqualified range against the active sheet and disclose that choice. Workbook mutations require explicit sheet targets, with the new-sheet and qualified-chart-source exceptions described in the tool pages.
- Confirm approvals queue and are attributable to the tab/action. Stop withdraws pending approvals. Local guidance/Files/integration configuration is not a workbook mutation; its own consent policy applies.
- Read-back, backups and Office.js execution have explicit limits. See [limits](foundations/limits.md), [backups](tools/backups.md) and the owning tool pages rather than retaining pre-fix numeric assumptions here.
- Verification checklists keep older Result columns as historical evidence. Current expectations follow merged contracts; integration-fixes logs retain the builds they actually tested.
- Real turns in this revision use GPT-6 Luna through the existing login. Do not refresh or replace the shared `PI_EXCEL_WEB_AUTH_JSON` grant.

## Working rules

- Commit with `docs: add agent-description/{path}` or `docs: revise agent-description/{path}`. This repo does not add AI attribution to commit messages.
- Do not modify anything outside `agent-description/`.
- Do not add files outside the README's structure without updating the structure and coverage table.
- When a behavior cannot be determined, write what you could determine, put the rest in "Open questions", and move on.
- Never renumber checklist IDs or triage IDs once used.
