# Glossary

The vocabulary used across these documents. When a document uses one of these words, it means exactly this. Terms the user-facing description already owns (turn, tab, busy, Stop, execution mode, backup) mean what [product-description/glossary.md](../product-description/glossary.md) says; this glossary adds the agent's side.

## The agent and its turn

**Agent.** The model running inside one Pi tab, together with what it is given: the system prompt, the tool schemas, the conversation history, the auto-context notes and the tool results. "The agent is told" means the text appears in a model request; nothing else reaches it. The agent cannot see tool cards, toasts, banners, the grid, or anything the user sees but Pi does not send.

**Turn.** As in [product-description](../product-description/glossary.md): everything from the user's message to idle. From the agent's side a turn is a loop of *model calls*, each of which may ask for *tool calls*.

**Model call.** One request to the model and its streamed answer. Every model call is built afresh: the system prompt, the tool schemas, the shaped history, and a newly gathered *auto-context* note. The product's "model response" is the answer to one model call.

**Tool call.** The agent's unit of interaction: one call to one tool with one set of arguments, from the moment the model emits it to the moment its *tool result* joins the history. Its five phases are *forms the call*, *rejected at once*, *touches Excel*, *while running*, *the result* (see the [README](README.md#document-template)). Several tool calls from one model call run in parallel unless one of them is marked sequential (none of the workbook tools are).

**Batch of calls.** All the tool calls emitted by one model call. They start together; their results return together, in call order, before the next model call.

## What the agent is told

**System prompt.** The text sent as the model's instructions: identity, rules, execution mode, active integrations, local services, skills, tools, workspace, workflow and conventions. It changes when those inputs change; earlier character counts are measurements of older builds, not size guarantees. See [what the agent is told](foundations/what-the-agent-is-told.md).

**Tool schema.** A tool's name, description and parameter schema, sent with model requests. The active configuration and registry determine the set; the older 26-tool/36,500-character measurement is not a fixed contract. Stable tool order and fingerprints avoid needless prefix changes during refresh.

**Auto-context.** Workbook structure, Files and other-people edit notes stored in the conversation when changed, plus fresh bounded selection context for each model request. Durable notes survive later tool calls and session restore; unchanged structure is not appended repeatedly. See [auto-context](foundations/auto-context.md).

**Blueprint.** The bounded workbook overview in auto-context: workbook, sheets and used ranges, row-1 previews, tables, objects and named ranges. Its signature is refreshed after structure changes, including user changes; it is appended when changed, not re-sent as missing on each call.

**Selection context.** Fresh per-request selection address, true dimensions and a bounded head/tail sample with row/column labels, formulas, errors and omitted counts. Whole-column selections are sampled, not loaded in full. See [limits](foundations/limits.md).

**Change list.** Durable conversation notes about other people's workbook edits, consumed per runtime rather than by the first tab to request context. Pi's own writes are excluded; added sheets receive listeners too.

**Rules and conventions.** The user's persistent instructions ("All my files", "This file") and formatting defaults, rendered into the system prompt. Both are empty in the default configuration.

## Tool results

**Tool result.** What a tool call returns to the agent: a list of content blocks (text, and for chart images an image) plus an `isError` flag. Only the content and the flag reach the model.

**Details.** Structured metadata attached to a tool result (`details.kind`, `changes`, `recovery`, `outputTruncation`). Details drive the tool card in the UI and are never sent to the model. A fact that is only in the details is a fact the agent does not know.

**Error-shaped result.** Readable failure text in a tool result. Workbook failures now set `isError: true`; the older false flag is a historical defect, not a contract. Contrast a *thrown error*, which the loop converts to an error result.

**Thrown error.** A failure raised out of the tool, which the agent loop turns into a result with `isError: true`. Only a few paths do this: schema validation, a declined Confirm-mode approval ("Mutation cancelled by user (Confirm mode)."), an abort before the call started ("Aborted"), and the unsupported-host wrapper.

**Blocked.** A mutation the tool refused before writing, with a result that says so ("⛔ **Write blocked** — …"): overwrite protection, or invalid formula syntax. Nothing in the workbook changed.

**Silent success.** A result that says the call did what was asked when the workbook shows otherwise: the values did not land, landed differently, or broke something outside the target. The most harmful result shape, because the agent has no reason to look again.

**False failure.** A result that says the call failed when the change was in fact applied. The agent may retry, give up, or tell the user nothing changed.

**Verification.** Checked read-back after a mutation, with a named full or sampled scope and explicit conversion, formula, validation and calculation limitations. It does not prove all dependent outputs or protect against a later co-author edit.

**Backup note.** A mutation's receipt naming the saved snapshot ID or stating the actual reason no backup was made. IDs are readable model content as well as UI metadata.

**Truncation.** Cutting a tool result's text to the per-result limit (50 KB or 2,000 lines by default) and appending `[Output truncated: …]`. Applied once, when the call finishes. See [limits](foundations/limits.md).

**History shaping.** Replacing eligible large answered-request tool results with bounded model-facing previews. Active unanswered work is preserved. The stored session retains the original text. See [context pressure](hostility/context-pressure.md) for selection and compaction boundaries.

**Compaction.** Replacing the older part of the history with a written summary, triggered automatically near the context limit or by `/compact`. After it, the agent sees a user message beginning "The conversation history before this point was compacted into the following summary:".

## Excel

**Excel batch.** One `Excel.run` session: a set of queued operations sent to Excel with `context.sync()`. A tool call makes one or more batches. Every request and every response of a batch is bounded by the *payload limit*. A batch has no timeout and cannot be cancelled once sent.

**Payload limit.** Excel's cap on the size of one request to it and one response from it (about 5 MB on Excel on the web). A batch that exceeds it fails with "The response payload size has exceeded the limit…" (or the request equivalent) and returns nothing. Operations queued before the oversized read in the same batch may already have run.

**Workbook lock.** The per-taskpane mutation queue across tabs. Reads wait for earlier registered mutations from their own runtime, not arbitrary other-runtime writes. It is not a co-authoring lock. See [the lock and parallel calls](hostility/the-lock-and-parallel-calls.md).

**Read call, mutating call.** The coordinator's classification by name/action. Workbook mutations acquire the lock after approval when required; same-runtime reads respect earlier writes. Local configuration uses its own consent policy instead of pretending to mutate the workbook.

**Structural call.** A mutating call whose impact is "structure": `modify_structure`, `execute_office_js`, chart create/delete, sheet hide/show. Only structural calls invalidate the blueprint.

**Active sheet.** The sheet the user is looking at in Excel at the moment a batch runs. Read calls can resolve an *unqualified address* against it; guarded mutations require explicit sheet targets.

**Qualified address, unqualified address.** `Sheet2!B3` names its sheet; `B3` does not. Reads can use the active sheet and disclose that resolution; guarded mutations refuse an omitted sheet. Sheet names with spaces or quotes are written `'My Sheet'!A1`; tools accept them with or without the quotes.

**Used range.** The rectangle from the first to the last non-empty or formatted cell of a sheet. It need not start at A1. Tools report its size ("200001 rows × 5 cols") but not its position.

**Raw value, displayed text.** What `Range.values` returns (a date as serial 45292, 5% as 0.05) versus what the cell shows (`Range.text`). Compact/detailed reads use displayed text and annotate hidden stored values; CSV and Python `values` remain raw, while Python `text` provides the display.

**Merged area.** A block of cells merged into one. Only its top-left cell holds a value; the others read as empty and silently ignore writes.

**Spill, spill anchor, spill child.** A dynamic-array formula (`=SEQUENCE(5)`) in its *anchor* cell fills a *spill range*; the other cells are *children*: they hold values but no formula. Writing into a child blocks the spill and turns the anchor into `#SPILL!`.

**Cell-edit mode.** The user is typing in a cell or the formula bar. Workbook tools wait up to 60 seconds for editing to end before a precise failure; context collection does not hold the model request for that wait. See [cell-edit mode](hostility/cell-edit-mode.md).

**Calculation mode.** Automatic or Manual (or Automatic except tables). Checked writes calculate written formulas in Manual mode and disclose that dependents elsewhere can remain stale. See [recalculation](hostility/recalculation.md).

**Workbook identity.** As in [product-description](../product-description/glossary.md): a hash of the workbook's address, or a document token where available without a location. Location-less tab layouts are document-scoped rather than sharing one global slot; file-specific rules still require usable identity.

## Hostility

**Hostility.** Anything about Excel's state or behavior that makes a correct tool call fail, return misleading text, or succeed in a way the agent cannot detect. The ten families are the rows of every "Hostility and interrupts" table; each has a document in [hostility/](hostility/).

**Escape hatch.** `execute_office_js`, used by the agent to learn what the structured tools do not tell it (displayed text, merge areas, protection, visible rows, calculation mode). In the real turns run for this project the agent reached for it in six of eleven turns. Heavy use of the escape hatch is a symptom of a gap in a structured tool.
