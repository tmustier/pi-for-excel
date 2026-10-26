# Glossary

The vocabulary used across these documents. When a document uses one of these words, it means exactly this. Terms the user-facing description already owns (turn, tab, busy, Stop, execution mode, backup) mean what [product-description/glossary.md](../product-description/glossary.md) says; this glossary adds the agent's side.

## The agent and its turn

**Agent.** The model running inside one Pi tab, together with what it is given: the system prompt, the tool schemas, the conversation history, the auto-context notes and the tool results. "The agent is told" means the text appears in a model request; nothing else reaches it. The agent cannot see tool cards, toasts, banners, the grid, or anything the user sees but Pi does not send.

**Turn.** As in [product-description](../product-description/glossary.md): everything from the user's message to idle. From the agent's side a turn is a loop of *model calls*, each of which may ask for *tool calls*.

**Model call.** One request to the model and its streamed answer. Every model call is built afresh: the system prompt, the tool schemas, the shaped history, and a newly gathered *auto-context* note. The product's "model response" is the answer to one model call.

**Tool call.** The agent's unit of interaction: one call to one tool with one set of arguments, from the moment the model emits it to the moment its *tool result* joins the history. Its five phases are *forms the call*, *rejected at once*, *touches Excel*, *while running*, *the result* (see the [README](README.md#document-template)). Several tool calls from one model call run in parallel unless one of them is marked sequential (none of the workbook tools are).

**Batch of calls.** All the tool calls emitted by one model call. They start together; their results return together, in call order, before the next model call.

## What the agent is told

**System prompt.** The fixed text sent as the model's instructions: identity, rules, execution mode, active integrations, local services, available skills, tools, workspace, workflow and conventions. About 13,700 characters in the default configuration. It changes only when one of its inputs changes (rules, execution mode, integrations, skills, conventions); see [what the agent is told](foundations/what-the-agent-is-told.md).

**Tool schema.** A tool's name, description and parameter schema, as sent with every model call. 26 tools are sent in the default configuration (about 36,500 characters of JSON). Every tool is sent on every call; nothing is disclosed progressively.

**Auto-context.** The note Pi builds before every model call and places, as an extra user message starting `[Auto-context]`, just before the latest real user message. It can hold the *blueprint*, the *workspace files summary*, the *selection context* and the *change list*. It is never stored in the session: the next model call gets a newly built one, and the old one is gone. See [auto-context](foundations/auto-context.md).

**Blueprint.** The workbook overview inside the auto-context: workbook name, each sheet with its used-range size, its row-1 "Headers", tables and object counts, and named ranges. It is the same text `get_workbook_overview` returns without a sheet. It is cached per workbook and rebuilt only after a structural tool call by Pi.

**Selection context.** The part of the auto-context that describes the user's current selection: its address, a window of the sheet around it (five rows above and below, up to 20 columns of the used range), the formulas inside the selection and any errors in the window.

**Change list.** The part of the auto-context headed `**User changes since last message:**`, listing cell addresses that changed since the list was last read. It is one list for the whole taskpane, emptied each time any model call reads it.

**Rules and conventions.** The user's persistent instructions ("All my files", "This file") and formatting defaults, rendered into the system prompt. Both are empty in the default configuration.

## Tool results

**Tool result.** What a tool call returns to the agent: a list of content blocks (text, and for chart images an image) plus an `isError` flag. Only the content and the flag reach the model.

**Details.** Structured metadata attached to a tool result (`details.kind`, `changes`, `recovery`, `outputTruncation`). Details drive the tool card in the UI and are never sent to the model. A fact that is only in the details is a fact the agent does not know.

**Error-shaped result.** A tool result whose text reports a failure ("Error writing cells: …", "Error reading …") but whose `isError` flag is false. Every workbook tool reports Excel failures this way; to the agent loop, the compaction summarizer and the history shaper, the call succeeded. Contrast a *thrown error*.

**Thrown error.** A failure raised out of the tool, which the agent loop turns into a result with `isError: true`. Only a few paths do this: schema validation, a declined Confirm-mode approval ("Mutation cancelled by user (Confirm mode)."), an abort before the call started ("Aborted"), and the unsupported-host wrapper.

**Blocked.** A mutation the tool refused before writing, with a result that says so ("⛔ **Write blocked** — …"): overwrite protection, or invalid formula syntax. Nothing in the workbook changed.

**Silent success.** A result that says the call did what was asked when the workbook shows otherwise: the values did not land, landed differently, or broke something outside the target. The most harmful result shape, because the agent has no reason to look again.

**False failure.** A result that says the call failed when the change was in fact applied. The agent may retry, give up, or tell the user nothing changed.

**Verification.** The read-back some mutating tools do after writing, and report as "**Verified values:**" or as a list of formula errors. It reads only the target range, once, immediately after the write.

**Backup note.** The sentence a mutating tool appends to its result about the automatic backup: nothing when a backup was made, otherwise "ℹ️ Backup not created …". The backup ID itself is only in the details.

**Truncation.** Cutting a tool result's text to the per-result limit (50 KB or 2,000 lines by default) and appending `[Output truncated: …]`. Applied once, when the call finishes. See [limits](foundations/limits.md).

**History shaping.** Rewriting older tool results before each model call: every tool result older than the six most recent that is longer than 1,200 characters (or holds an image) is replaced, for the model only, by a 500-character preview headed `[Compacted tool result]`. The stored session keeps the full text. See [context pressure](hostility/context-pressure.md).

**Compaction.** Replacing the older part of the history with a written summary, triggered automatically near the context limit or by `/compact`. After it, the agent sees a user message beginning "The conversation history before this point was compacted into the following summary:".

## Excel

**Excel batch.** One `Excel.run` session: a set of queued operations sent to Excel with `context.sync()`. A tool call makes one or more batches. Every request and every response of a batch is bounded by the *payload limit*. A batch has no timeout and cannot be cancelled once sent.

**Payload limit.** Excel's cap on the size of one request to it and one response from it (about 5 MB on Excel on the web). A batch that exceeds it fails with "The response payload size has exceeded the limit…" (or the request equivalent) and returns nothing. Operations queued before the oversized read in the same batch may already have run.

**Workbook lock.** The per-taskpane queue that runs Pi's mutating tool calls one at a time, across all tabs. Read calls never wait for it. See [the lock and parallel calls](hostility/the-lock-and-parallel-calls.md).

**Read call, mutating call.** The coordinator's classification of a tool call, from its name and `action`. Read calls run at once; mutating calls take the workbook lock, are subject to Confirm-mode approval, and may trigger a blueprint refresh. Unknown tools count as mutating.

**Structural call.** A mutating call whose impact is "structure": `modify_structure`, `execute_office_js`, chart create/delete, sheet hide/show. Only structural calls invalidate the blueprint.

**Active sheet.** The sheet the user is looking at in Excel at the moment a batch runs. An *unqualified address* is resolved against it.

**Qualified address, unqualified address.** `Sheet2!B3` names its sheet; `B3` does not and lands on the active sheet. Sheet names with spaces or quotes are written `'My Sheet'!A1`; tools accept them with or without the quotes.

**Used range.** The rectangle from the first to the last non-empty or formatted cell of a sheet. It need not start at A1. Tools report its size ("200001 rows × 5 cols") but not its position.

**Raw value, displayed text.** What `Range.values` returns (a date as the serial 45292, 5% as 0.05, £1,234.50 as 1234.5) versus what the cell shows (`Range.text`). The structured tools send raw values; displayed text reaches the agent only through `execute_office_js`.

**Merged area.** A block of cells merged into one. Only its top-left cell holds a value; the others read as empty and silently ignore writes.

**Spill, spill anchor, spill child.** A dynamic-array formula (`=SEQUENCE(5)`) in its *anchor* cell fills a *spill range*; the other cells are *children*: they hold values but no formula. Writing into a child blocks the spill and turns the anchor into `#SPILL!`.

**Cell-edit mode.** The state Excel is in while the user is typing in a cell or the formula bar and has not pressed Enter, Tab or Escape. On Excel on the web every batch fails at once with "Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again." See [cell-edit mode](hostility/cell-edit-mode.md).

**Calculation mode.** Automatic or Manual (or Automatic except tables). In Manual, formulas written by Pi show stale results (often 0) until something recalculates. See [recalculation](hostility/recalculation.md).

**Workbook identity.** As in [product-description](../product-description/glossary.md): a hash of the workbook's address, or unknown. Backups, rules for "This file" and the blueprint cache are keyed by it; when it is unknown they share one "unknown" key.

## Hostility

**Hostility.** Anything about Excel's state or behavior that makes a correct tool call fail, return misleading text, or succeed in a way the agent cannot detect. The ten families are the rows of every "Hostility and interrupts" table; each has a document in [hostility/](hostility/).

**Escape hatch.** `execute_office_js`, used by the agent to learn what the structured tools do not tell it (displayed text, merge areas, protection, visible rows, calculation mode). In the real turns run for this project the agent reached for it in six of eleven turns. Heavy use of the escape hatch is a symptom of a gap in a structured tool.
