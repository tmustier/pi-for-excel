# Cell-edit mode

## Summary

Excel may reject requests while the user types in a cell. Requested workbook tools wait with backoff for up to 60 seconds, with the activity "Waiting for you to finish editing a cell", then proceed once editing ends. Auto-context does not hold the model request for this wait; it reports unreadable selection instead.

## What the agent receives

If the user leaves edit mode in time, the normal tool result. If not, a precise error says the user is typing, how long Pi waited, whether an earlier step reached Excel, and to ask the user to finish before retrying. It does not claim nothing changed when an earlier sync succeeded.

## What the agent does

The tool waits rather than immediately abandoning the task. After the bounded failure the model can ask the user to finish editing. A Stop during the wait cancels it, preventing delayed execution after the user later presses Enter.

## By phase of a tool call

| Phase | Effect |
| --- | --- |
| Forms the call | Fresh selection can report edit mode. |
| Rejected at once | Non-workbook validation can still reject. |
| Touches Excel | Retry starts at 250 ms and backs off to 2 seconds, within the 60-second budget. |
| While running | Retry only if no earlier sync succeeded. A partially applied operation is never replayed automatically. |
| Result | Success after editing ends, explicit edit-mode failure, or cancellation. |

## Per tool

Workbook tools using the shared Excel wrapper follow this rule. Tools that do not touch Excel do not wait. A multi-step operation can still have earlier work applied before a later step encounters edit mode; inspect its result before retrying.

## Host differences

The retry handles Excel's edit-mode rejection on web and desktop. A host may itself queue requests; its individual run still has a timeout. Windows native behavior was not exercised here.

## Proposed fix

The shared bounded wait is merged; this heading remains for historical links, not as a planned feature.

## Open questions and verification

Code reviewed in `src/excel/helpers.ts` at `1f51cf8`. Historical immediate-failure timings belong to the old verification pass; real platform timing is not re-measured here.

Verified against pi-for-excel commit `1f51cf8`.
