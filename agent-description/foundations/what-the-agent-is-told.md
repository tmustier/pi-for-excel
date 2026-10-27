# What the agent is told

## Summary

Everything the agent knows comes from five sources, rebuilt into every model call: the system prompt, the tool schemas, the stored conversation (shaped), the [auto-context](auto-context.md) note, and the tool results it receives. This document owns the first two and the rules about when they change. It also records what the agent is *not* told about Excel, because most of the hostility in this project is a gap between what the prompt promises and what Excel does.

The active tool configuration has a stable ordered prefix, but the merged prompt/schemas include paged reads, workbook-state reporting and checked-write limitations. Exact tool counts and character sizes from the 26 October capture are historical measurements, not the size of the current contract.

## The simple case

A user opens Pi on a workbook and asks a question. The model call carries, in order:

1. The system prompt, as the provider's instructions.
2. The 26 tool schemas, all of them, every time.
3. The stored conversation: earlier user messages, the agent's answers and tool calls, and tool results (older large ones [shaped](../hostility/context-pressure.md) to previews).
4. The auto-context note, just before the newest user message.
5. The newest user message.

The agent answers, perhaps with tool calls; their results are appended and the next model call is built the same way, with a freshly gathered auto-context note.

## The system prompt, section by section

The sections appear in this order. Sections marked *conditional* are omitted when empty.

| Section | What it tells the agent (quoted or summarized) |
| --- | --- |
| Identity | "You are Pi, an AI agent embedded in Microsoft Excel as a sidebar add-in. You can read, modify, format, and research — working directly in the user's live workbook." |
| `## Rules` | How to maintain persistent rules with the `instructions` tool; user rules may be updated freely, workbook rules only after showing the exact text and getting confirmation; then the current rules under "All my files" and "This file", "(No rules set.)" by default. |
| `## Execution mode` | "Current mode: **Auto**", with "Favor low-friction execution for workbook mutations. Do not add extra pre-execution confirmation prompts beyond existing safety gates." In Confirm: "Ask for explicit user confirmation before mutating workbook tools. Treat destructive structure operations as high-risk and reconfirm before proceeding." |
| `## Active Integrations` *(conditional)* | Web Search by default: use `web_search` then `fetch_page`, cite sources, and a warning that queries leave the machine. |
| `## Connections` *(conditional)* | Status of tools that need a connection; never ask for keys in chat. Absent by default. |
| `## Local Services` *(conditional)* | The native Python bridge and tmux bridge, both "not running" by default, with the skills to read if the user needs them. |
| `## Available Agent Skills` | Five bundled skills (`extending-pi`, `mcp-gateway`, `python-bridge`, `tmux-bridge`, `web-search`) in an XML list, with "read each skill once per session". |
| `## Tools` | One line per core tool from the capability table, then `extensions_manager`, `execute_office_js`, `execute_wps_js`; a `### Python` subsection (Pyodide by default); notes on the Files workspace, `assistant-docs/`, and the rule to use `execute_office_js` "for workbook features not covered by structured tools (for example Excel tables with filters and PivotTables)"; Office.js hygiene ("call `context.sync()` after `load()`") and the warning that browser globals trigger an approval even in Auto. |
| `## Workspace` | Folder conventions (`notes/`, `workbooks/<name>/`, `scratch/`, `imports/`, `assistant-docs/`) and the memory contract. |
| `## Workflow` | Nine rules: read first; edit scope; verify writes ("write_cells auto-verifies and reports errors"); overwrite protection ("Ask the user before setting allow_overwrite=true"); prefer formulas; report changes; plan complex tasks; analysis is read-only; extension requests. |
| `## Conventions` | A1 notation with examples (`"A1:D10"`, `"Sheet2!B3"`); cite cells; Arial 10; "For large ranges, read a sample first to understand the structure."; tables with headers; cell styles and presets; font colour conventions (blue inputs, black formulas, green cross-sheet links). |
| Convention overrides *(conditional)* | Custom presets and non-default conventions, when set. |

## The tool schemas

Every model call carries every tool, in a fixed order: the 17 core tools (`get_workbook_overview`, `read_range`, `write_cells`, `fill_formula`, `search_workbook`, `modify_structure`, `format_cells`, `conditional_format`, `charts`, `trace_dependencies`, `explain_formula`, `view_settings`, `comments`, `instructions`, `conventions`, `workbook_history`, `skills`), then `tmux`, `python_run`, `libreoffice_convert`, `python_transform_range`, `files`, `execute_office_js`, `extensions_manager`, and the Web Search integration's `web_search` and `fetch_page`. `tmux` and `libreoffice_convert` are listed although their bridges are not running; calling them returns a gate result that points to a skill.

The largest schemas are `conventions` (5,661 characters), `format_cells` (4,072), `charts` (2,294), `tmux` (2,262) and `conditional_format` (2,038). There is code for choosing a smaller tool bundle from the user's wording ("format", "chart", "comment"), but it is not used: the tool list is always the full list so that the prompt cache is not split by intent.

Arguments are validated against the schema before a tool runs. A failure becomes a [thrown error](the-tool-contract.md#thrown-errors) the agent can read: "Validation failed for tool "{name}": - {path}: {message} … Received arguments: …".

## When it changes

The system prompt and tool list are rebuilt when the runtime's capabilities are refreshed: a change of execution mode, rules, conventions, integrations, connections, skills or extension tools. A changed prompt is appended to the conversation as a new system message rather than replacing the old one; the latest one wins. The tool list is replaced only when its fingerprint changes. Nothing about the workbook (sheets, selection, edits) is in the system prompt; that is all in the auto-context, which keeps the prompt's prefix stable for caching.

## What the agent is not told

The original information gaps and their current handling are:

- **Raw versus displayed values.** Compact/detailed reads now show displayed text and hidden stored values where relevant; CSV is raw. Python input includes values and text. See [workbook representation](the-workbook-as-the-agent-sees-it.md).
- **Size.** Reads page with continuation, writes chunk, and results name verification scope and uncertain outcomes; the agent must still follow continuation and obey [limits](limits.md).
- **Workbook state.** Structured reads report merges, spills, protection, filters, validation and calculation state, with unavailable inspection named. They are snapshots, not a promise state cannot change.
- **Verification has limits.** Checked writers name full/sampled scope and calculate written formulas in Manual mode, but do not prove all dependents elsewhere are fresh. See [recalculation](../hostility/recalculation.md).
- **The user keeps working.** Nothing says the user may be typing in a cell, switching sheets, or editing the cells the agent just read. See [cell-edit mode](../hostility/cell-edit-mode.md) and [the user moves or edits](../hostility/the-user-moves-or-edits.md).
- **Parallel calls.** Same-runtime reads wait for previously registered mutations; they do not wait for other runtimes' future writes. See [the lock and parallel calls](../hostility/the-lock-and-parallel-calls.md).
- **Explicit mutation targets.** Guarded workbook mutations require a named sheet rather than following the user's active sheet. Read calls can still resolve unqualified addresses against the active sheet and disclose that resolution. See [workbook representation](the-workbook-as-the-agent-sees-it.md).
- **Errors.** Workbook failures are text plus `isError`, with definite-refusal versus possibly-applied wording; not every failure means nothing changed.
- **Older results are shaped.** Eligible results from answered requests can become bounded previews; active unanswered work is retained. The model is not given the original result merely because storage still has it.

## Hostility and interrupts

The system prompt and schemas are fixed for the duration of a model call; none of the hostility families changes them. Changing the execution mode mid-turn appends a new system prompt that applies from the next model call; Confirm-mode approvals already open are unaffected. Compaction leaves the system prompt in place and replaces history before it with a summary.

## Interactions with other systems

**What the agent is told.** This document.

**Verification.** The prompt's rule 3 tells the agent to trust `write_cells` verification and to re-read downstream outputs; see [the tool contract](the-tool-contract.md#verification) for what verification covers.

**Backups and recovery.** The prompt lists which tools make automatic backups (in `workbook_history`'s line) and the execution-mode section mentions "fail-closed restore safeguards". It does not say backups are skipped above 20,000 cells.

**The workbook lock and parallel calls.** Not mentioned.

**Execution mode and approvals.** The execution-mode section changes with the mode and asks for an up-front plan in Confirm mode. The `execute_office_js` note says browser globals trigger an approval in any mode.

**Truncation and history shaping.** Not mentioned in the prompt; each truncated result carries its own notice.

**Auto-context.** Not described in the prompt. The agent learns its shape only from the notes themselves.

**Prompt caching.** The prompt and tool list are kept byte-stable across calls (no timestamps, fixed tool order) so providers can cache the prefix; volatile state goes into the auto-context at the tail.

**Host differences.** The prompt is the same on desktop and web. `execute_wps_js` is described but registered only on WPS.

**What the user sees.** Rules and conventions are edited by the user in `/rules` and Settings; see [product-description](../../product-description/README.md).

## Edge cases

- The prompt's Python section says "Two Python tools are always available" and that Pyodide works out of the box; on a host without WebAssembly workers the tools return an error at call time.
- The prompt names `execute_wps_js` on every host; on Excel the tool does not exist, and calling it returns "Tool execute_wps_js not found" as a thrown error.
- A rules change during a turn takes effect at the next model call, inside the same turn.
- The "Local Services" section is probed at session start and is not refreshed when the user starts a bridge mid-session (read from code).

## Open questions and verification

- Confirmed: the section order, sizes and tool list, from the captured request of a real turn on 2026-10-26 (TOLD-01 to TOLD-03).
- Read from code: when the prompt is rebuilt; the unused disclosure bundles.
- The historical information gaps above have been revised against merged prompt/read/write contracts. Overwrite confirmation policy still follows the prompt; this pass does not implement proposed product calls. Exact prompt sizes and every provider serialization were not recaptured.

Verified against pi-for-excel commit `1f51cf8` (source inspection).
