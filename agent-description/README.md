# Pi for Excel agent description

A written description of Pi for Excel from the standpoint of the agent inside it: what the model is told, what each tool call does to the workbook, and exactly what comes back, including what happens when Excel gets in the way halfway through.

## Purpose

[`product-description/`](../product-description/README.md) describes Pi for Excel as the user meets it. This project describes the same product as the agent meets it. The agent is the model that receives the system prompt, the tool schemas, the auto-context notes and the tool results, and that acts on the workbook only through tool calls. From its side, Pi for Excel is a large state chart too: it moves through it with tool calls, and Excel answers with text. Most of that behavior is defined implicitly, spread across the system prompt (`src/prompt/`), the tools (`src/tools/`), the context injector (`src/taskpane/context-injection.ts`, `src/context/`), the message conversion and history shaping (`src/messages/`), compaction (`src/compaction/`, `src/commands/builtins/export.ts`) and the workbook coordinator (`src/workbook/`). There is no single place that says "when the agent calls X, this is what Excel does, this is what the agent is told, and this is what happens if the user does Y while it runs."

This project is that place. It exists because Excel can be hostile to an agent in several ways that are not yet solved: the user can be typing in a cell, a range can be too big to send, a sheet can be protected, a formula can be waiting for a recalculation that never comes, and the agent is often told none of this. The documents find those problems by asking the same questions of every tool, and [bug-triage.md](bug-triage.md) ranks them by how much they hurt task success.

The documents are for people who change the agent's experience: whoever edits the system prompt, a tool, a result format or the context pipeline. They are written from the agent's side. They describe what the agent receives and what its calls do, not how the code is organized.

### What this is not

- Not a description of the user's experience. What the user sees (tool cards, approval dialogs, banners) belongs to [product-description/](../product-description/README.md); these documents link there instead of restating it.
- Not API documentation. Tool behavior rules live in [src/tools/DECISIONS.md](../src/tools/DECISIONS.md); this project describes their consequences for the agent.
- Not a prompt-engineering guide. It records what the agent is told and what it receives; it does not prescribe wording, except where a proposed fix in [bug-triage.md](bug-triage.md) needs it.
- Not a description of what the model says. Model behavior varies; where a real model turn was run, the document reports what that model did as an observation, never as a rule.

## Conventions

- Describe the agent's experience, not the code. "The agent is told the cells were written, and three of the four values are gone" rather than "`targetRange.values = padded` is not checked against the merge areas".
- Quote what the agent receives exactly, in code spans or fenced blocks: tool result text, auto-context notes, error messages. The exact text is the experience.
- Technical detail goes in block quotes, prefixed with `Technical note:`. Use it only when the mechanism changes what the agent would expect.
- Use sentence case for headings.
- Name the vocabulary consistently. The [glossary](glossary.md) is the source of truth for terms like *tool call*, *tool result*, *model call*, *auto-context*, *blueprint*, *change list*, *error-shaped result*, *silent success*, *false failure*, *payload limit* and *cell-edit mode*.
- Every document ends with the pi-for-excel commit it was verified against and a list of open questions.
- When a behavior is surprising, say so and say why it is that way if the reason is known. Do not smooth it over.
- Mark how each claim is known: "confirmed" means observed in Excel on the web in this project's probe workbook (see [verification/README.md](verification/README.md)); "read from code" means not observed.

## The work to be done

Each document describes one tool, one tool family, one foundation, or one form of Excel hostility, in full, including its edge cases and its interactions with the rest of the agent's world.

### Document template

Every tool document follows the same skeleton so that documents are comparable and nothing is skipped. The full skeleton is in [goal.md](goal.md#document-skeleton).

1. **Summary.** What the tool lets the agent do, and how it is described to the agent.
2. **The simple case.** The common call and what comes back.
3. **The call, event by event.** The five phases of a *tool call*: **forms the call** (what the agent knows when it decides, what the schema accepts); **rejected at once** (schema validation, argument checks, a gate or policy, a declined approval, overwrite protection: nothing in the workbook changes); **touches Excel** (the first Excel batch; for mutations the workbook lock and the backup snapshot; from here the workbook may change); **while running** (batches and syncs, waiting for the lock, what else can happen meanwhile); **the result** (exactly what text the agent gets, what is truncated, what is verified, what is only in details, how the result ages in history). A Mermaid `stateDiagram-v2` of the states the call passes through.
4. **Workbook and host state.** The variant axis: a table with the same rows in every document, saying what each state does when it is true before the call and when it changes while the call runs:
    - Sheet protection
    - Calculation mode
    - Merged cells
    - Tables, filters and hidden rows
    - Data validation
    - Dynamic arrays and spills
    - External links
    - Very large used range
    - Number formats and locale
    - Host and API set (desktop or web)
    - Execution mode (Auto or Confirm)
5. **Hostility and interrupts.** The same checklist in every document, in this order:
    - The user is editing a cell
    - The user changes the selection or the active sheet
    - The user changes data the call depends on
    - Another tab's write holds the workbook lock
    - Excel throttles, times out or rejects the payload
    - Recalculation is pending
    - The workbook is co-authored
    - Autosave or a save boundary
    - The user presses Stop
    - The context fills: shaping, compaction or overflow
6. **Interactions with other systems.** In this fixed order: **What the agent is told.** **Verification.** **Backups and recovery.** **The workbook lock and parallel calls.** **Execution mode and approvals.** **Truncation and history shaping.** **Auto-context.** **Prompt caching.** **Host differences.** **What the user sees.**
7. **Edge cases.**
8. **Open questions and verification.** The commit, what was confirmed and how, what is read from code only, and what looks like a bug.

Item 5 matters most. Asking the same hostility questions of every tool is how the gaps show up.

### Method

For each document:

1. Read the tool (`src/tools/{tool}.ts`), its wrappers (`with-workbook-coordinator.ts`, `output-truncation.ts`, `execution-policy.ts`), and what the agent is told about it (`src/tools/capabilities.ts`, the schema descriptions, `src/prompt/system-prompt.ts`).
2. Read the matching tests in `tests/`: `workbook-write-inspect-undo-contract.test.ts`, `tools-output-truncation.test.ts`, `context-injection.test.ts`, `execution-policy.test.ts`, `tools-workbook-coordinator.test.ts`, `with-workbook-coordinator.test.ts`, `search-workbook-tool.test.ts` and the per-tool tests.
3. Draft the document.
4. Probe the real tool against Excel on the web: `node scripts/orb/excel-web-session.mjs --check` opens Pi in the orb's own workbook, and the taskpane is a CDP target on port 9223 in which the tools can be imported from the dev server (`await import("/src/tools/registry.ts")`) and called directly, wrapped exactly as the runtime wraps them. Build the hostile state first (protect a sheet, merge cells, filter a table, switch to manual calculation, type in a cell without pressing Enter) and record the result text verbatim. Use real model turns (`--prompt`, or the composer driven over CDP with the model requests captured from the network) sparingly: they share one subscription's rate limit.
5. Record the commit verified against.

### Verification

The `verification/` directory holds one checklist per cluster of documents, each item a single observable claim with setup, steps, expected result, a priority, and what it needs. A tester runs them, records `pass`, `fail`, or `blocked` in the Result column, and files every failure in [bug-triage.md](bug-triage.md). A document moves from `drafted` to `verified` only when every P1 and P2 item for it has passed or been filed. The first pass, run while drafting, is recorded in [verification/README.md](verification/README.md#results-so-far).

### Order of work

1. **Pilot: reading a range.** The most common call, small enough to settle the template, and already hostile (raw values, payload limits, merged cells, filters).
2. **Foundations: what the agent is told, auto-context, the tool contract, the workbook as the agent sees it, limits.**
3. **Writing.** `write_cells` and `fill_formula`, where the worst failures are.
4. **The hostility documents**, one per interrupt family, which own the detail every tool document links to.
5. **The other tools**, verification checklists, then the triage.

### Scope decisions

- **Surface.** The agent of one Pi tab in the default configuration: Auto execution mode, no extensions, no experimental flags, no rules or conventions set, the bundled integrations as shipped (Web Search active), one provider. Observed on Excel on the web (ExcelApi up to 1.19) with GPT-6.1 Sol (500k context window); desktop behavior is described from code and marked as such.
- **Unit of interaction.** A *tool call*, with the five phases named in the template above. A turn is many tool calls; what a turn is belongs to [product-description/foundations/turns.md](../product-description/foundations/turns.md). The auto-context, the system prompt and history shaping are not tool calls but are described in the foundations because every call depends on them.
- **Variant axis.** The state of the workbook and the host, not modifier keys. The rows are fixed as listed above.
- **Interrupt list.** The ten Excel hostility families above. They are derived from the five families in the method (the user's explicit abort: Stop; the user doing something else mid-way: editing a cell, moving the selection, changing data; the environment failing: throttling, payload limits, timeouts; something else changing the target: another tab, co-authors, recalculation, autosave; the channel changing: the context filling and history being rewritten). Each family has one [hostility/](hostility/) document that owns its detail.
- **Cross-cutting order.** Fixed as listed above. **What the user sees** is always last and always a link into `product-description/`.
- **WPS, extensions, bridges.** Out of scope. WPS has its own tool implementations; extension tools are defined by the extension; the tmux and native Python bridges are off unless the user runs them. `python_run` and `python_transform_range` run in-browser (Pyodide) by default and are in scope.
- **Provider failures.** Retries, rate limits and overflow banners are the product's concern ([product-description/](../product-description/README.md)); this project covers only what the agent receives when they happen (a failed turn, a compacted history).
- **Numbered rules.** Prose documents, not numbered specifications. Stable heading anchors are enough for cross-references.
- **Location.** This project lives in `agent-description/` at the repository root, next to `product-description/`, for the same reason: files under `docs/` are bundled into the taskpane as assistant docs.

## Structure

```
README.md                          this file
goal.md                            the standing instructions for whoever drafts
AGENTS.md, CLAUDE.md               entry points for agents: read README.md, then goal.md
glossary.md                        shared vocabulary
bug-triage.md                      agent-experience problems ranked by harm to task success,
                                   with causes in code and proposed fixes

verification/
  README.md                        how to probe the tools in Excel on the web; results so far
  foundations-and-reading.md       checklists for foundations/ and the reading tools
  writing.md                       checklists for the writing and other mutating tools
  hostility.md                     checklists for hostility/

foundations/
  what-the-agent-is-told.md        the system prompt, the tool schemas, rules and conventions
  auto-context.md                  the blueprint, the selection, the change list and the files
                                   summary injected before every model call
  the-tool-contract.md             what a tool result is: content and details, error-shaped
                                   results, verification, backups notes, truncation
  the-workbook-as-the-agent-sees-it.md
                                   addresses, sheets, the active sheet, raw values versus
                                   displayed text, identity
  limits.md                        every number that bounds a call: payload limits, truncation,
                                   history shaping, compaction, timeouts

tools/
  reading-a-range.md               the pilot: read_range in compact, csv and detailed modes
  the-workbook-overview.md         get_workbook_overview, workbook and sheet level
  searching.md                     search_workbook
  tracing-and-explaining.md        trace_dependencies and explain_formula
  writing-cells.md                 write_cells
  filling-a-formula.md             fill_formula
  formatting.md                    format_cells and conditional_format
  structure-and-view.md            modify_structure and view_settings
  charts-and-comments.md           charts and comments
  office-js.md                     execute_office_js, the escape hatch
  python.md                        python_run and python_transform_range
  backups.md                       workbook_history
  instructions-conventions-and-skills.md
                                   instructions, conventions, skills and files

hostility/
  cell-edit-mode.md                the user is typing in a cell
  the-user-moves-or-edits.md       selection, active sheet and data changing under a call
  the-lock-and-parallel-calls.md   the workbook lock, other tabs, parallel calls, racing approvals
  payload-limits-and-large-ranges.md
                                   the request and response payload limits and big sheets
  throttling-and-timeouts.md       slow, throttled and never-answering Excel calls
  recalculation.md                 manual calculation and pending recalculation
  co-authoring-and-autosave.md     other people's edits, autosave and the save boundary
  stop.md                          Stop at each phase of a tool call
  context-pressure.md              truncation, history shaping, compaction and overflow
```

## Coverage

Status is one of `not started`, `drafted`, or `verified`.

| Document | Status |
| --- | --- |
| glossary.md | drafted |
| bug-triage.md | drafted |
| verification/ (3 checklists) | drafted |
| foundations/what-the-agent-is-told.md | drafted |
| foundations/auto-context.md | drafted |
| foundations/the-tool-contract.md | drafted |
| foundations/the-workbook-as-the-agent-sees-it.md | drafted |
| foundations/limits.md | drafted |
| tools/reading-a-range.md | drafted |
| tools/the-workbook-overview.md | drafted |
| tools/searching.md | drafted |
| tools/tracing-and-explaining.md | drafted |
| tools/writing-cells.md | drafted |
| tools/filling-a-formula.md | drafted |
| tools/formatting.md | drafted |
| tools/structure-and-view.md | drafted |
| tools/charts-and-comments.md | drafted |
| tools/office-js.md | drafted |
| tools/python.md | drafted |
| tools/backups.md | drafted |
| tools/instructions-conventions-and-skills.md | drafted |
| hostility/cell-edit-mode.md | drafted |
| hostility/the-user-moves-or-edits.md | drafted |
| hostility/the-lock-and-parallel-calls.md | drafted |
| hostility/payload-limits-and-large-ranges.md | drafted |
| hostility/throttling-and-timeouts.md | drafted |
| hostility/recalculation.md | drafted |
| hostility/co-authoring-and-autosave.md | drafted |
| hostility/stop.md | drafted |
| hostility/context-pressure.md | drafted |

## Reference

The source of truth is the pi-for-excel repository itself (this repo), at the commit each document cites. The relevant locations are:

- `src/prompt/system-prompt.ts`, `src/tools/capabilities.ts`: what the agent is told
- `src/tools/registry.ts`, `src/tools/*.ts`: the tools, their schemas and their result text
- `src/tools/with-workbook-coordinator.ts`, `src/tools/execution-policy.ts`, `src/workbook/coordinator.ts`: read/mutate classification, the workbook lock, Confirm-mode approvals
- `src/tools/output-truncation.ts`, `src/context/window-budgets.ts`: per-result truncation
- `src/taskpane/context-injection.ts`, `src/taskpane/context-refresh-decision.ts`, `src/context/`: auto-context (blueprint, selection, change list)
- `src/messages/convert-to-llm.ts`, `src/messages/tool-result-shaping.ts`: history shaping
- `src/compaction/`, `src/commands/builtins/export.ts` (`runCompactCommand`): compaction
- `src/workbook/recovery-log.ts`, `src/tools/mutation/`, `src/tools/recovery-metadata.ts`: automatic backups and their notes
- `src/excel/helpers.ts`, `src/utils/format.ts`: address parsing and the text rendering of values
- `src/taskpane/init.ts` (`buildRuntimeCapabilities`, the `Agent` construction): how it is all wired
- `node_modules/@earendil-works/pi-agent-core/dist/agent-loop.js`: the agent loop (argument validation, parallel tool execution, `transformContext` per model call)
- `tests/`: behavioral tests
