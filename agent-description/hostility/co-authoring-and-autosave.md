# Co-authoring and autosave

## Summary

On Excel on the web, and in desktop workbooks stored on OneDrive or SharePoint, other people can edit the same workbook while Pi works, and the workbook autosaves continuously. The second pass exercised co-authoring on the web with a second guest editor: the same anonymous folder-link workbook opened in a fresh browser profile (CDP port 9224), typing into the grid with real keystrokes while Pi's tools ran in the first profile. Pi's model of the sheet drifts silently: co-authors' edits are labelled as the user's, a co-author's later Enter overwrites a write Pi has just verified, and a co-author's row insertion moves Pi's write away from the address in its result ([B-23](../bug-triage.md#b-23-a-co-authors-edits-overwrite-or-move-pis-verified-writes-with-no-notice)).

## Co-authors' edits

- **Remote edits are reported as the user's.** A co-author's edit to `Calc!Z3` arrived through `onChanged` as `{source: "Remote", triggerSource: "Unknown"}`. The change list does not look at `source` (confirmed for the event; the tracker's labelling is [B-08](../bug-triage.md#b-08-the-change-list-reports-pis-own-writes-misses-new-sheets-and-lasts-one-model-call)).
- **Co-authoring changes how Pi's own writes are reported.** Before the co-author joined, Pi's writes arrived with `triggerSource: "ThisLocalAddin"`. From the moment the co-author joined they arrived with `"Unknown"`, the same as a user's edit, and stayed that way after the co-author left (confirmed). A fix that filters on `triggerSource` alone would stop working in exactly this case.
- **Last writer wins, after verification.** The co-author typed in `Co!B4` without pressing Enter. Pi's `write_cells` of `555` there succeeded and reported `Verified values: 555`. When the co-author pressed Enter, `B4` became `9`. Nothing told the agent (confirmed).
- **A co-author's edit mode does not block Pi.** Pi's reads and writes succeeded while the co-author was mid-edit; only the local user's edit mode causes [cell-edit-mode](cell-edit-mode.md) failures (confirmed).
- **Structure moves under the agent.** The co-author inserted a row above row 2 while Pi wrote `[["Total", "=SUM(B2:B5)"]]` at `Co!A6`. The result said `Written to **Co!A6:B6**` with `836` verified; the cells ended in row 7 as `=SUM(B3:B6)`. Every address the agent read before is now one row off, and the blueprint is not refreshed (confirmed).
- **Sync delay.** Remote values reached Pi within about 4 s. A sheet Pi added took over 20 s to appear in the co-author's tabs (confirmed).
- Overwrite protection checks only at the moment of the write. A co-author's edit between the agent's read and its write is overwritten if `allow_overwrite` is set.
- Restoring a backup overwrites co-authors' later edits to those cells.
- Workbook rules ("This file") are stored on this machine, not in the workbook. A co-author's Pi does not see them.

## Autosave and the save boundary

Automatic backups are cleared when the workbook goes from unsaved to saved. On the web with autosave, the product-description drafts report that `isDirty` is always false and backups "may be cleared within seconds". In this session, backups made between 17:55 and 18:23 were all still listed by `workbook_history list` (confirmed), with workbook identity unknown (anonymous-link workbook). So the clearing was not reproduced here; it may depend on identity being known. The agent is never told backups were cleared. It finds out only when a restore ID is missing.

## Open questions and verification

- Confirmed: COAUTH-01 to COAUTH-05 (second guest profile, web).
- The co-author's session once dropped to a Microsoft sign-in page after a reload; reopening the folder link restored it. Guests share one anonymous identity, so both editors appeared as the same "GC" avatar.
- Backup clearing on the web with a known identity was not tested (needs a signed-in workbook).

Verified against pi-for-excel commit `ec453df`.
