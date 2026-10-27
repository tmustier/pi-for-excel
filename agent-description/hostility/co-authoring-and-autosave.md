# Co-authoring and autosave

## Summary

Other people can edit a cloud workbook while Pi works and Excel autosaves. At the merged baseline Remote edits have distinct attribution and changed structure/edit notes persist in conversation. These are observations, not co-author locks: a later edit can overwrite or move a value after Pi verified it. The historical second-guest pass below establishes those race windows, not current attribution or backup clearing.

## Co-authors' edits

- **Remote edits are distinguished.** Excel's Remote source is labelled as other people's edits, including when triggerSource is Unknown. Historical `Calc!Z3` events established this combination.
- **Co-authoring changes how Pi's own writes are reported.** Before the co-author joined, Pi's writes arrived with `triggerSource: "ThisLocalAddin"`. From the moment the co-author joined they arrived with `"Unknown"`, the same as a user's edit, and stayed that way after the co-author left (confirmed). A fix that filters on `triggerSource` alone would stop working in exactly this case.
- **Last writer wins, after verification.** The co-author typed in `Co!B4` without pressing Enter. Pi's `write_cells` of `555` there succeeded and reported `Verified values: 555`. When the co-author pressed Enter, `B4` became `9`. Nothing told the agent (confirmed).
- **A co-author's edit mode does not block Pi.** Pi's reads and writes succeeded while the co-author was mid-edit; only the local user's edit mode causes [cell-edit-mode](cell-edit-mode.md) failures (confirmed).
- **Structure moves under the agent.** The co-author inserted a row above row 2 while Pi wrote `[["Total", "=SUM(B2:B5)"]]` at `Co!A6`. The result said `Written to **Co!A6:B6**` with `836` verified; the cells ended in row 7 as `=SUM(B3:B6)`. Every address the agent read before is now one row off, and the blueprint is not refreshed (confirmed).
- **Sync delay.** Remote values reached Pi within about 4 s. A sheet Pi added took over 20 s to appear in the co-author's tabs (confirmed).
- Overwrite protection checks only at the moment of the write. A co-author's edit between the agent's read and its write is overwritten if `allow_overwrite` is set.
- Restoring a range backup compares against Pi's recorded after-state. Later changes require Keep my edits or Restore anyway; the agent tool refuses with `later_edits` unless explicitly overridden. It is still not an atomic co-author lock.
- Workbook rules ("This file") are stored on this machine, not in the workbook. A co-author's Pi does not see them.

## Autosave and the save boundary

Saving/autosave and reloading do not clear automatic backups. A first identity transition migrates document-token backups to the saved identity. Retention and explicit deletion remove entries; atomic storage updates preserve other panes' appends. Sessions shared across panes save a conflict copy rather than overwriting a newer revision. Neither mechanism coordinates live co-author cell edits.

## Open questions and verification

- Confirmed: COAUTH-01 to COAUTH-05 (second guest profile, web).
- The co-author's session once dropped to a Microsoft sign-in page after a reload; reopening the folder link restored it. Guests share one anonymous identity, so both editors appeared as the same "GC" avatar.
- This pass inspected the merged tracker/recovery paths; it did not repeat the second-guest race experiment against them.

Verified against pi-for-excel commit `1f51cf8` (source inspection; COAUTH observations are historical).
