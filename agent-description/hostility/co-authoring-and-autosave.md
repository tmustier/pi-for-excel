# Co-authoring and autosave

## Summary

On Excel on the web, and in desktop workbooks stored on OneDrive or SharePoint, other people can edit the same workbook while Pi works, and the workbook autosaves continuously. Neither was exercised with a second person in this pass. This document records what the code implies for the agent.

## Co-authors' edits

- Values and structure change under the agent with no notice. Excel reports remote changes through `onChanged` with source `Remote`. The change list does not distinguish them and, like all its entries, labels them as the user's (read from code).
- The blueprint does not refresh for a co-author's new sheets.
- Overwrite protection checks only at the moment of the write. A co-author's edit between the agent's read and its write is overwritten if `allow_overwrite` is set.
- Restoring a backup overwrites co-authors' later edits to those cells.
- Workbook rules ("This file") are stored on this machine, not in the workbook. A co-author's Pi does not see them.

## Autosave and the save boundary

Automatic backups are cleared when the workbook goes from unsaved to saved. On the web with autosave, the product-description drafts report that `isDirty` is always false and backups "may be cleared within seconds". In this session, backups made between 17:55 and 18:23 were all still listed by `workbook_history list` (confirmed), with workbook identity unknown (anonymous-link workbook). So the clearing was not reproduced here; it may depend on identity being known. The agent is never told backups were cleared. It finds out only when a restore ID is missing.

## Open questions and verification

- No second user was available; nothing about co-authoring was observed.
- Backup clearing on the web with a known identity was not tested.

Verified against pi-for-excel commit `ec453df`.
