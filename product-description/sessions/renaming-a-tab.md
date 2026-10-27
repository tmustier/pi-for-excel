# Renaming a tab

## Summary

As of [PR #805](https://github.com/tmustier/pi-for-excel/pull/805) (CL-08, CL-12), a tab shows its conversation title, whether written by the model or chosen by the user. Without a title it shows its stable "Chat N" fallback. A user-chosen title always wins over automatic generation.

Rename by double-clicking a tab, choosing "Rename tab…" in its context menu, or typing `/name Q3 budget review`. All three work while Pi is busy: changing a label does not change the running conversation. No provider is needed for manual renaming.

## Automatic titles

After the first successful exchange containing user text and an assistant answer, Pi makes one small model request for a 3–5 word title in the user's language. This is a separate request with an isolated session/cache namespace, no conversation tools, and bounded excerpts of the first exchange. It does not alter the main agent's system prompt, tool schemas or history.

The attempt is persisted before the request, so reload or resume does not repeat a failed paid attempt. A failed or empty result leaves "Chat N". A title returned after the user renamed the tab cannot overwrite that choice. Closing the runtime cancels its pending title request.

The saved title is used consistently in the tab strip, resume picker, close toast and recently closed list. Existing saved titles also display in tabs; they are not replaced merely because they predate model-written titles.

## The simple case

Type `/name Q3 budget review` to replace the model-written title with a chosen title.

## Manual rename

- Double-click targets that tab; "Rename tab…" targets the right-clicked tab even if another tab is active. `/name` targets the active tab.
- Dialog paths use a translated in-taskpane text dialog, not `window.prompt`, on desktop and web. Its current-label prefill is selected; it traps focus and returns focus to the opener when dismissed.
- Confirming trims the text and saves the chosen title. Cancel or Escape changes nothing. If the target runtime disappeared while the dialog was open, nothing is renamed. Tab shortcuts do not act behind the modal dialog.
- Clearing the dialog resets the title to "Chat N". `/name` without text shows its usage toast instead of clearing a title.
- Successful paths use "Renamed to {title}"; clearing uses "Tab name reset". Names and chosen-title metadata survive reload, resume and reopen. Session saving begins with the first user message, not the first answer.

### Ends at once

Cancel, Escape, a missing target or `/name` without text records no title change. Being busy is not a refusal.

## Interactions with other systems

Renaming adds no conversation message, sends no model request and changes no workbook cells or backups. Different sessions may have identical names. Duplicating a tab creates a separate session with its own chosen copy title; later renaming the source does not rename the copy. Model switching stays in the same tab and does not invent a fork title.

Long labels truncate visually; the full title remains available in the tooltip and accessible tab name. UI text is translated; generated titles follow the conversation language. User titles are not translated. There is no rename history or rename Undo.

## Open questions and verification

PR #805 includes public session-persistence regression tests for isolated title generation, chosen-title races and failed-title non-retry on resume. Real Excel on the web acceptance is recorded on that PR; desktop release acceptance remains separate. The earlier native-prompt/reload failures are historical observations in the [checklist](../verification/sessions-and-commands.md#sessionsrenaming-a-tabmd), not the current contract.
