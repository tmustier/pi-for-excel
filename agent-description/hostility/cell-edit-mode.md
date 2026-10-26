# Cell-edit mode

## Summary

While the user is typing in a cell or the formula bar and has not pressed Enter, Tab or Escape, Excel on the web refuses every Office.js batch. Every tool that touches Excel fails at once, and auto-context loses the selection. It is the most common hostility in normal use, because users keep working while Pi runs.

## What the agent receives

Every Excel batch fails in 3–12 ms with (confirmed):

```
Excel is in cell-editing mode. Please exit the edit mode by pressing ENTER or TAB or selecting another cell, and then try again.
```

wrapped by the tool: `Error reading "Dates!A1:B2": …`, `Error writing cells: …`, `Error executing Office.js: …`. Each is an [error-shaped result](../foundations/the-tool-contract.md#error-shaped-results). The auto-context note for the next model call simply has no selection context. It does not say why. Change-tracker registration also fails if the taskpane starts while a cell is being edited (confirmed: `onChanged.add` threw), and nothing retries it.

## What the agent does

In a real turn ("read Big!A1:E30, then Dates!A1:D4, then write the sum into Calc!D1"), the user started typing after the first read. The second read failed, and the agent ended the turn without writing: "Reading Dates!A1:D4 failed because Excel is in cell-editing mode. Stopped; Calc!D1 was not changed. Press Enter or Tab to finish editing, then I can continue." That is a sensible reaction to what it was told, but the task is lost: the user has to notice, press Enter, and ask again.

## By phase of a tool call

| Phase | Effect |
| --- | --- |
| Forms the call | The agent cannot know edit mode is on; nothing in its context says so. |
| Rejected at once | Not applicable; Pi checks nothing before calling Excel. |
| Touches Excel | Fails immediately. Nothing changes in the workbook. |
| While running | A multi-batch tool (`write_cells`: pre-read, write, read-back) can succeed on early batches and fail on later ones. If the write landed and the read-back failed, the result is a false failure (read from code). |
| The result | The edit-mode error string. |

## Per tool

Every Excel-touching tool behaves the same way. `instructions`, `conventions`, `skills`, `files`, `workbook_history list`, `web_search` and `python_run` are unaffected.

## Host differences

Desktop Excel is documented to queue add-in requests until edit mode ends rather than fail them. That was not observed here. If true, desktop calls hang (with no timeout, see [throttling and timeouts](throttling-and-timeouts.md)) where web calls fail.

## Proposed fix

See [B-06](../bug-triage.md#b-06-cell-edit-mode-fails-every-call-at-once-and-ends-the-task): retry in `excelRun` with backoff for up to about 30 s on `InvalidOperationInCellEditMode`, then return a result that says the user is editing and nothing changed. Have the auto-context say so as well.

## Open questions and verification

- Confirmed on Excel on the web: EDIT-01 to EDIT-04 and one real turn.
- Desktop behavior not observed.

Verified against pi-for-excel commit `ec453df`.
