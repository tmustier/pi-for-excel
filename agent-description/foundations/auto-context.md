# Auto-context

## Summary

Before every model call, Pi gathers what it knows about the workbook and the user's activity and sends it to the agent as an extra user message beginning `[Auto-context]`. It can have four parts, always in this order: the [blueprint](#the-blueprint) (the workbook overview), the [workspace files summary](#the-workspace-files-summary), the [selection context](#the-selection-context), and the [change list](#the-change-list). It is how the agent knows, without a tool call, which sheets exist, what the user has selected, and what the user has edited.

The note is built afresh for every model call, including each call inside a turn after tool results come back, and it is never stored in the session. The next call does not see the previous note; it sees a new one. That single fact explains most of this document's surprises.

## The simple case

The user selects `Dates!A1:D4` and asks a question. The agent's request contains, just before the user's message (confirmed, abridged):

```
[Auto-context]
[Workbook context refresh: initial]

## Workbook: orb-98fb59459bd2-20261026175207.xlsx

### Sheets (10)
1. **Sheet1** — 1 rows × 1 cols
2. **Dates** — 4 rows × 5 cols
   Headers: Date, Growth, Price, Code
...

**Selection:** Dates!A1:D4
**Context:** Dates!A1:G6

| Date | Growth | Price | Code |  |  |  |
| --- | --- | --- | --- | --- | --- | --- |
| 45292 | 0.05 | 1234.5 | 00123 | abc |  |  |
...

**User changes since last message:** Sheet1: G6
```

The agent answers. If it calls tools, the next model call carries a new note: the blueprint again (now labelled "context missing from history"), the selection as it is now, and no change list unless more edits arrived in between.

## Where the note goes and how long it lives

The note is inserted immediately before the latest real user message in the conversation. During a turn's tool loop that is still the message that started the turn, so the note sits before the user's request and before all of the turn's tool calls and results. On the next turn it moves to just before the new message; the old note is gone.

> Technical note: the note is produced by the agent's `transformContext` hook (`src/taskpane/context-injection.ts`). pi-agent-core calls it before every model call and uses its output only for that request; the agent's stored messages never contain it.

Each part is gathered with a time limit, and any part that fails or times out is silently left out: workbook identity 1.2 s, blueprint 2.5 s, files summary 1.2 s, selection 1.5 s. The agent is never told a part is missing or why.

## The blueprint

The blueprint is the same text `get_workbook_overview` returns with no `sheet` argument (see [the workbook overview](../tools/the-workbook-overview.md)): the workbook name, then each sheet with its position, name, visibility if not Visible, used-range size ("200001 rows × 5 cols", or "empty"), up to eight non-empty values from row 1 as "Headers", each table's name and size, object counts, and the workbook's visible named ranges.

It is headed by a reason line:

| Reason line | When |
| --- | --- |
| `[Workbook context refresh: initial]` | The first model call of the tab. |
| `[Workbook context refresh: context missing from history]` | Every later model call (confirmed). |
| `[Workbook context refresh: workbook structure changed]` | The first call after Pi made a [structural call](../glossary.md#excel). |
| `[Workbook context refresh: workbook switched]` | The workbook identity changed. |

The "context missing from history" line appears on every call because the injector checks the stored messages for an earlier note, and there never is one. The intent was to send the blueprint once and again only when it changed; the effect is that it is sent every time. That is harmless for correctness and costs its length on every call.

The blueprint is cached per workbook and rebuilt only after Pi's own structural calls (`modify_structure`, `execute_office_js`, chart create or delete, sheet hide or show). It is *not* rebuilt when the user adds, renames or deletes a sheet, when another tab or co-author does, or when Pi's `write_cells` grows a sheet or writes a new row 1. A sheet the user added stayed out of the blueprint (confirmed). When the workbook identity is unknown, all workbooks share one cache entry.

What the blueprint does not contain: where each used range starts (a sheet whose data is at `B2:G6` reads "5 rows × 6 cols"), which row is really the header (row 1 is used whatever it holds; a sheet whose row 1 is a spill showed "Headers: #SPILL!, #SPILL!"), protection, merged areas, filters, hidden rows, data validation, the calculation mode, or charts' and pivots' locations.

## The workspace files summary

`[Workspace files refresh: initial]` or `[Workspace files refresh: files changed]`, followed by a `### Workspace` list of notes, imports and workbook-tagged files. It is sent on the first call only when there are relevant files, and afterwards only when the set changes. Unlike the blueprint it is sent once: after its first appearance the injector remembers its signature in memory, so later calls skip it. Because the note is not stored, a model call after that first one carries no files summary at all (read from code). With no relevant files, the summary is never sent.

## The selection context

```
**Selection:** Dates!G7
**Context:** Dates!A2:G6

| 45292 | 0.05 | 1234.5 | 00123 | abc |  |  |
| --- | --- | --- | --- | --- | --- | --- |
| 45323 | 0.125 | 99 | 00456 | xyz |  |  |
...
**Selected formulas:** …
⚠️ **Errors nearby:** …
```

The window is the selection plus five rows above and below, clipped to the sheet's used range, across the used range's first 20 columns. It is rendered as a plain markdown table whose first row is treated as the table header; there are no row numbers and no column letters, so the agent must work out addresses from the "Context:" line. Values are raw, as everywhere (see [the workbook as the agent sees it](the-workbook-as-the-agent-sees-it.md)). Formulas are listed only for cells inside the selection; errors for the whole window.

Three behaviors are hostile:

- **No size limit.** The window grows with the selection. With `Big!A1:A5000` selected, the note was 311,761 characters and was sent on every model call of the turn; the context meter went from 3% to 32% of a 500k window for "Reply with just OK." (confirmed). A 20,000-row selection measures 873,849 characters, more than a 200k-token window holds.
- **Whole columns vanish.** Selecting a whole column, row or sheet makes the read of the selection's own values exceed the payload limit, and the selection context is silently omitted — even on a four-row sheet (confirmed for `B:B`, `A:XFD`, and `Dates!A:A`).
- **Outside the used range.** A selection below or right of the data still shows the used-range window, which then does not contain the selected cell (confirmed: selection `Sheet1!C3`, context `Sheet1!B2`).

When the user is typing in a cell, the selection read fails and the selection context is omitted for that call (confirmed).

## The change list

```
**User changes since last message:** Sheet1: G6
```

Pi listens for cell changes on every sheet that was visible when the taskpane opened. Each change adds an address; before each model call the list is deduplicated, grouped by sheet (ten addresses per sheet, then "… (+N more)"), sent, and emptied. At most 50 changes are kept between reads.

Four faults make it unreliable:

- **Pi's own writes count as the user's.** Excel reports Pi's writes with `triggerSource: "ThisLocalAddin"` and the user's with `"Unknown"`; the list does not look. A probe write to `Sheet1!H1` was reported as `**User changes since last message:** Sheet1: H1` (confirmed).
- **New sheets are never watched.** Sheets added after the taskpane opened, by the user or by Pi, are not tracked; an edit to `Dates!G6` (Dates was added after opening) never appeared (confirmed).
- **One model call only.** The list is emptied by the call that reads it, and the note is not stored, so the second model call of a turn no longer knows about the edits. In a real turn the change line was in request 1 of 6; answering in request 6, the agent said "I wasn't given an edit log" (confirmed).
- **One list for all tabs.** The list belongs to the taskpane, not the tab; whichever tab makes the next model call consumes it (read from code).

The list carries addresses only, not values, and not who changed them.

## Hostility and interrupts

| Event | Effect on the note |
| --- | --- |
| The user is editing a cell | Selection context omitted; blueprint served from cache if it was cached, otherwise omitted; change list unaffected. |
| The user changes the selection or the active sheet | The next model call's note shows the new selection; the agent is not told it moved. |
| The user changes data the call depends on | Added to the change list if the sheet is watched; shown at the next model call only. |
| Another tab's write holds the workbook lock | No effect; gathering the note takes no lock. |
| Excel throttles, times out or rejects the payload | The affected part is omitted after its time limit; whole-column selections always hit the payload limit. |
| Recalculation is pending | Values in the selection window are the last calculated ones. |
| The workbook is co-authored | Co-authors' edits are reported by Excel as remote changes; whether they reach the change list was not observed. The blueprint does not refresh for their structural changes. |
| Autosave or a save boundary | No effect. |
| The user presses Stop | No effect on the note; no further model call is made. |
| The context fills | The note is not counted in Pi's own token estimate until a model answers; a large selection inflates every request and can overflow the window, and compaction cannot shrink it because it is rebuilt each time. |

## Interactions with other systems

**What the agent is told.** The system prompt never describes the note; the agent learns its meaning from the headings.

**Verification.** No interaction.

**Backups and recovery.** No interaction.

**The workbook lock and parallel calls.** Gathering takes no lock and runs before the model call, so it cannot overlap the agent's own tool calls; it can overlap another tab's writes.

**Execution mode and approvals.** No interaction.

**Truncation and history shaping.** The note is not a tool result: it is never truncated and never shaped.

**Auto-context.** This document.

**Prompt caching.** The note sits before the turn's first user message, so everything after it in the turn (all tool calls and results) follows a block that changes whenever the selection, the blueprint or the change list changes. A change between two model calls of the same turn invalidates the cached tail.

**Host differences.** The payload limits that remove whole-column selections are those of Excel on the web; desktop was not observed.

**What the user sees.** Nothing; the note is never displayed. See [product-description/turn/auto-context.md](../../product-description/README.md) (planned there).

## Edge cases

- With no workbook (the taskpane opened outside Excel), every part fails and no note is sent.
- An empty sheet's selection is `**Selection:** Sheet1!A1 (sheet is empty)`.
- A selection on a hidden sheet cannot happen; a selection on a sheet whose used range is narrower than the selection shows only the used columns.
- Fifty changes or more are reduced to the last fifty before deduplication.
- The very first model call of a session restored from storage gets a blueprint headed "initial" again.

## Open questions and verification

- Confirmed on Excel on the web, 2026-10-26, from captured requests: placement, the reason lines, re-sending, the selection window and its size, whole-column omission, the change-list faults (own writes, new sheets, one call), omission during cell-edit mode.
- Read from code: the per-part time limits; the files summary being sent once only; the shared list across tabs; co-author changes.
- Looks like a bug: [B-04](../bug-triage.md#b-04-the-selection-context-has-no-size-limit-and-disappears-for-whole-column-selections), [B-08](../bug-triage.md#b-08-the-change-list-reports-pis-own-writes-misses-new-sheets-and-lasts-one-model-call), [B-12](../bug-triage.md#b-12-the-blueprint-goes-stale-and-is-re-sent-on-every-model-call), [B-21](../bug-triage.md#b-21-small-result-format-slips).

Verified against pi-for-excel commit `ec453df`.
