# What the agent is told

## Summary

Everything the agent knows comes from five sources, rebuilt into every model call: the system prompt, the tool schemas, the stored conversation (shaped), the [auto-context](auto-context.md) note, and the tool results it receives. This document owns the first two and the rules about when they change. It also records what the agent is *not* told about Excel, because most of the hostility in this project is a gap between what the prompt promises and what Excel does.

In the default configuration (captured from the request on 2026-10-26) the system prompt is 13,675 characters and the 26 tool schemas are 36,557 characters of JSON. Both are sent with every model call, unchanged unless one of their inputs changes.

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

Read against the hostility found in this project, the prompt and schemas are silent on:

- **Values are raw.** Dates arrive as serial numbers, percentages as fractions, currency without symbols; zero arrives as `0.0000e+0`. Nothing says so, and nothing says `detailed` mode shows formats. See [the workbook as the agent sees it](the-workbook-as-the-agent-sees-it.md).
- **Size.** "Read a sample first" is the only guidance. Nothing says a whole column cannot be read, that about 10,000 rows × 5 columns is near the ceiling on the web, or that a large write can fail after applying. See [payload limits](../hostility/payload-limits-and-large-ranges.md).
- **Merged cells, protection, filters, spills, data validation, calculation mode.** None is mentioned. The only hint is that tables with filters and PivotTables need `execute_office_js`.
- **Verification has limits.** "write_cells auto-verifies and reports errors" is true only for the written cells, and only when calculation is automatic. See [recalculation](../hostility/recalculation.md).
- **The user keeps working.** Nothing says the user may be typing in a cell, switching sheets, or editing the cells the agent just read. See [cell-edit mode](../hostility/cell-edit-mode.md) and [the user moves or edits](../hostility/the-user-moves-or-edits.md).
- **Calls in one batch are unordered.** Nothing says a read issued alongside a write may see the old value. See [the lock and parallel calls](../hostility/the-lock-and-parallel-calls.md).
- **Unqualified addresses follow the active sheet.** The conventions show qualified and unqualified examples side by side; the schemas say "If no sheet is specified, uses the active sheet" without saying the active sheet is the user's, at the moment of the call.
- **Errors arrive as text.** Nothing says tool failures are reported in the text of a normal result.
- **Older results are shaped.** Nothing says a result read six calls ago is now a 500-character preview.

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
- Looks like a gap: the hazards listed under [What the agent is not told](#what-the-agent-is-not-told) ([B-11](../bug-triage.md#b-11-the-agent-cannot-see-workbook-state-that-changes-what-its-calls-do)); rule 4's confirmation round trip on explicit requests ([B-18](../bug-triage.md#b-18-overwrite-protection-forces-a-confirmation-round-trip-on-explicit-requests)).

Verified against pi-for-excel commit `ec453df`.
