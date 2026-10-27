# Auto-context

## Summary

Auto-context has two lifetimes. Workbook structure, relevant Files and other people's edits become durable conversation notes when changed. The current selection is read afresh, bounded and appended at the end of every model request without being saved. The agent should treat samples as samples and read omitted cells with tools.

## The simple case

A first request includes a stored note beginning "Workbook structure (initial). Sizes are as of this note; a new note follows when sheets, tables, names, rows or columns change." The current selection follows the request history, headed "Current selection (re-read for every request, not stored in the conversation):". Later calls keep the structure note in history and receive a fresh selection, rather than resending an initial blueprint ahead of the prompt.

## Where the note goes and how long it lives

Durable notes are inserted before new user prompts or after tool results mid-turn. They are saved with the session, remain in subsequent requests and can be summarized by compaction. The transient selection is always trailing and is never saved. Source failures are logged; missing context is not evidence that a sheet or edit does not exist.

Source timeouts are 1.2 seconds for workbook identity and Files, 2.5 seconds for blueprint and 2 seconds for selection.

## The blueprint

Workbook name, sheets with visibility/used sizes/headers, tables, object counts and visible names. It refreshes initially, on changed structure or workbook identity, and when compaction removed the earlier note. Sheet and structure events invalidate cached context, including sheets added outside Pi. A size is a snapshot, not a guarantee of current values.

This chat's own structure changes defer the new blueprint until the next turn: its tool result already states the change and should not be mistaken for a new request to repeat it.

## The workspace files summary

Relevant notes, imports and workbook-tagged files produce a stored note initially and on a changed summary. Later calls retain it in history. Emptying a previously nonempty relevant set also produces a change note. No relevant files initially means no initial Files note.

## The selection context

Exact address, row/column counts and used range precede a labelled table of displayed values. The sample has a 240-cell budget, at most 20 columns and 4,000 table characters. Tall windows show 15 head rows and 5 tail rows; windows up to 22 rows can fit whole. Omitted rows and columns are stated. Selected formulas and nearby errors have their own bounded lists.

Whole rows/columns/sheets are clipped to used data. A selection outside used data is explicitly described as empty. An unreadable selection, including cell-edit mode, is reported instead of silently pretending nothing was selected. Auto-context does not wait through edit mode. It is not a substitute for `read_range`.

## The change list

The shared bounded edit log is read through a separate cursor per runtime. A request from another tab cannot consume this tab's changes. Notes exclude this chat's own writes, distinguish Excel user edits, remote co-author edits and Pi outside this chat, and remain in history on later calls. New sheets are watched and renames use current names. Addresses are reported, not complete cell values.

## Hostility and interrupts

| Event | Effect |
| --- | --- |
| Editing a cell | Selection reports unavailable; requested tools have a separate bounded edit wait. |
| Selection moves | Next request gets a fresh trailing selection. |
| Data changes | Other edits become a durable attributed note. |
| Another runtime holds the write lock | Context reads are not co-authoring isolation and may overlap it. |
| Excel rejects a source or times out | That source is omitted or reports unavailable, not an unbounded wait. |
| Recalculation pending | Selection values may be stale; explicit read reports workbook state. |
| Co-authoring | Remote edits are labelled separately; there is no snapshot isolation. |
| Save | Notes and backups are not cleared by save. |
| Stop | No further model request; already stored notes remain. |
| Context pressure | Bounded selection cannot grow with a million-row selection; durable notes can be compacted. |

## Interactions with other systems

**What the agent is told.** Headings explain lifetime and sampling. **Verification.** Auto-context is not a verified tool read. **Backups and recovery.** No mutations. **The workbook lock and parallel calls.** No cross-runtime isolation. **Execution mode and approvals.** No approval. **Truncation and history shaping.** Selection has its own budgets; durable notes participate in compaction. **Auto-context.** This document. **Prompt caching.** Durable notes become cached history; moving selection changes only the trailing request message. **Host differences.** API failures can reduce available context. **What the user sees.** No ordinary bubble; see [auto-context](../../product-description/turn/auto-context.md).

## Edge cases

- Bounded edit history is not an audit log and can lose older events.
- A source timing out does not mean the workbook is empty.
- Compaction may cause current structure or Files to be reintroduced when their stored note is absent.

## Open questions and verification

Code reviewed at `1f51cf8`. A real GPT-6 Luna turn in Excel on the web confirmed stored structure, an outside-chat edit note and exclusion of its own formula write from subsequent context. Remote co-author and desktop events were not exercised here. Historical pre-fix observations remain in verification logs.

Verified against pi-for-excel commit `1f51cf8`.
