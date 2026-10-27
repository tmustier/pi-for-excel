# Integration verification: workbook

This file reruns every item in [../workbook.md](../workbook.md) against the combined fix build, branch `integration-fixes`. That branch is `main` plus PRs #786–#799 (see `INTEGRATION.md` on the branch). The first full pass ran at `a2b3f2d`. After the branch was force-pushed to `e48d83c`, every item whose code changed was rerun there: fill_formula, the lock and Stop items (coordinator), the mode popover, Escape on Backups, the approval dialog's focus, refused format_cells and the backup-failure note. The rows give the result at `e48d83c` when the item was rerun there, and at `a2b3f2d` otherwise; the code behind those items did not change between the two builds. Both runs were on 27 October 2026, by an agent in an Amp orb.

Expected behavior is the checklist as amended by the fix PRs and `goal.md` "Fixes in flight". Where a PR deliberately changed behavior, the PR wins, and the note says "changed (#NNN)".

How it ran:

- **web-anon**: Pi in Excel on the web through `node scripts/orb/excel-web-session.mjs --check` (host Excel, platform OfficeOnline, check ok). The workbook was a fresh anonymous-link upload, and the dev server served this checkout (confirmed by fetching changed source files from the taskpane). The taskpane was driven over CDP on its own target: `Input.dispatchMouseEvent` clicks at element centres, `Input.dispatchKeyEvent` keys, and DOM and Office.js reads. Grid input (cell editing, Ctrl+S, Ctrl+Z/Y) went to the Excel page target.
- **Techniques**: the workbook lock was made observable by wrapping `Excel.run` in the taskpane with a 3–20 s delay, either for every call or only while a matching tool card existed. Reloads used CDP `Page.reload` on the taskpane target. Two panes were the same workbook opened in a second browser tab (`Target.createTarget`) with Pi opened there too. The full Excel reload was `Page.reload` on the OneDrive page. A backup-save failure was simulated by making the recovery log's `append` reject.
- **Model turns** ran on GPT-6 Luna at Min (GPT-6.1 Luna is not in the registry). The shared `PI_EXCEL_WEB_AUTH_JSON` login was never logged out or refreshed.

**Result: 115 items. 106 pass, 0 fail, 9 blocked (8 need desktop Excel, 1 needs a workbook with an identity).**

- Every item that failed on `main` and can run on the web now passes or behaves as its owning PR describes: WB-08, 13, 14, 21, 25; READ-10, 16; WRITE-10; APPROVE-03, 06, 13, 14, 17; REVERT-15, 16, 21, 22, 24, 26; BACKUP-06, 07. The other three `main` fails (BACKUP-08, 10, 11) are desktop downloads and are blocked here.
- No item that passed on `main` regressed.
- The pass found two new bugs; both are still open on `e48d83c` (see below).

## Results

### foundations/the-workbook.md

| ID | Result | Note |
| --- | --- | --- |
| WB-01 | pass | Confirm mode: "Read Sheet1!A1:A3" completed with no dialog; backup count unchanged (2). |
| WB-02 | pass | "Edit Sheet1!A1:A3 — 3 changed", no dialog, A1:A3 = 1, 2, 3. |
| WB-03 | pass | Backup "write_cells · Sheet1!A1:A3" listed. The result now ends "Backup saved: {full id}." (#794). |
| WB-04 | pass | The dialog opened before the write; D1 stayed empty. |
| WB-05 | pass, changed (#794) | Escape gives "The user declined this workbook change in the approval dialog (Confirm mode)." Pi replied; D1 stayed empty. |
| WB-06 | pass, changed (#793) | The badge is now "Waiting" (was "Locking…"). With `Excel.run` delayed 4 s, Chat 1 showed "Waiting" and the status bar "Waiting for workbook" while Chat 2 wrote. All 12 cells were written. Same at `e48d83c`. |
| WB-07 | pass | The close button was disabled, titled "Wait for workbook changes to complete", for the whole write window. In 100 ms samples it did not flicker back to enabled as it did on `main`. Same at `e48d83c`. |
| WB-08 | pass, fixed (#794) | The writing tab never showed "Waiting" on itself, and its close button stayed disabled throughout: Chat 2 from 5.3 to 30.6 s, then Chat 1 from 30.6 to 55.8 s (`e48d83c`). |
| WB-09 | pass, changed (#793) | The status bar shows "Changing workbook" while the active tab holds the lock, and "Waiting for workbook" while it waits. |
| WB-10 | pass, changed (#791, #792, #789) | 6 parallel writes with `Excel.run` delayed 3 s; Escape once all 6 cards existed. AN1 (in Excel): "Stopped by the user while Excel was still working on this change. The outcome is unknown…". AN2–AN6: "Stopped before running; nothing changed in the workbook." All cards red; AN1:AN6 empty; transcript "Request aborted" (`e48d83c`). |
| WB-11 | pass | `isDirty` was `false` after a write. |
| WB-12 | pass | Click on the grid, Ctrl+S, wait 9 s: 19 backups before and after. |
| WB-13 | pass, fixed (#794) | 19 backups before CDP `Page.reload` and 19 after 12 s. Tabs and mode were kept. |
| WB-14 | pass, fixed (#794) | The first backup on the fresh web-anon workbook was still listed after 10 s. |
| WB-15 | pass | The R1 backup taken after the reload was still listed after 10 s. |
| WB-16 | blocked | Desktop only. |
| WB-17 | pass | Pi wrote R1 = 5. On the grid, Ctrl+Z emptied R1 and Ctrl+Y put 5 back. |
| WB-18 | blocked | Desktop only. |
| WB-19 | pass, changed (#792) | With C10 held in edit mode, the dock showed "Waiting for you to finish editing a cell / Press Enter, Tab or Esc in Excel and Pi will continue". After about 58 s the card read: "Error reading "Sheet1!A1:A3": Excel is in cell-edit mode … Waited 58 s for the user to finish. Nothing was read or changed. Ask the user to finish editing the cell, then retry." Pi asked the user to finish editing. |
| WB-20 | pass | See BACKUP-01. |
| WB-21 | pass, fixed (#794) | Same workbook in two browser tabs. Opening pane B did not clear backups (28 before, 28 after). A wrote V1 and B wrote V2 at the same time: both panes listed both (30). B deleted V2, then A wrote V4: V2 stayed deleted in both panes (31). |
| WB-22 | blocked | Desktop only. |
| WB-23 | pass, changed (#791) | "Filled Sheet1!AG1:AG21000 — 21000 changed, no backup", with "ℹ️ No backup: this change covers 21,000 cells, more than the 20,000 a backup can hold. … was empty before, so clearing it undoes the change." and a partial-verification note. At `e48d83c` (one `copyFrom` paste) the turn took 5 s, and AG21000 = `=ROW()*2` = 42000. |
| WB-24 | pass | "Set freeze rows (Sheet1) — no backup", "ℹ️ Backup not created. …" |
| WB-25 | pass, fixed (#794) | See APPROVE-17. |
| WB-26 | pass | After closing pane B, a full Excel page reload and reopening Pi: URL "", workbook id null, the same document token (`c1ae1e3a…`). The 31 backups and the tabs were also kept (#794, #795). |

### workbook/reading.md

| ID | Result | Note |
| --- | --- | --- |
| READ-01 | pass | "Overview" card: "Sheets (1) Sheet1 — 22 rows × 30 cols", headers. |
| READ-02 | pass | "Read Sheet1!AF1:AF3", "All cells are empty." |
| READ-03 | pass | "No matches for "zzz" in any sheet (searched values)." |
| READ-04 | pass | "Trace precedents Sheet1!F1"; tree "2 precedents =A1+D1" with A1 (1) and D1 (5). |
| READ-05 | pass | "Explain formula Sheet1!F1", "Direct references (2)", "Sheet1!A1 → 1", "Sheet1!D1 → 5". |
| READ-06 | pass | The detailed read lists "Formulas F1: =A1+D1". |
| READ-07 | pass | A real click on the "Sheet1!B1:B2" link selected Sheet1!B1:B2; the card stayed collapsed. |
| READ-08 | pass | Confirm mode: the read did not ask (WB-01); only the write did (WB-04). |
| READ-09 | pass | See WB-19. |
| READ-10 | pass, fixed (#793) | The failed read card has `data-state="error"` (red), like failed writes. |
| READ-11 | pass, changed (#787) | `read_range AA1:AA5000` now returns one 500-row page with "Next page: read_range "Sheet1!AA501:AA1000"". A page over 50 KB (Big!A1:E500, 40-character text) is still cut: "[Output truncated: showing first 212 of 504 lines (50.0KB limit) …; full output saved to Files workspace: .tool-output/…]". |
| READ-12 | pass | "zzz" on a hidden sheet: a search without a sheet name finds nothing; a search on that sheet finds "Big!G1: zzz". |
| READ-13 | pass, changed (#792) | Read held 12 s; Escape at 4.7 s; the turn was idle at 5.4 s. Card: "Stopped by the user before Excel answered." (red), then "Request aborted". No workbook change. |
| READ-14 | pass | Chat 2's write was held 20 s from 2.0 s; Chat 1's read completed at 5.4 s (`e48d83c`). |
| READ-15 | pass | A second link clicked 140 ms later was ignored (selection B1:B2); one 630 ms later was honored (C1:C2). |
| READ-16 | pass, fixed (#787) | Used range A1:AD21000 (AA1:AA21000 `=ROW()`). A search for 20999 without a sheet name returned "Sheet1!AA20999: 20999 ← =ROW()". |

### workbook/writing.md

| ID | Result | Note |
| --- | --- | --- |
| WRITE-01 | pass | "CHANGES (3)", Before/After, "Sheet1!AJ1 ∅ 7". |
| WRITE-02 | pass | Drawer: "This operation changed 3 cells. Summary: … Citations: … Explanation uses a bounded metadata sample." |
| WRITE-03 | pass | "Write Sheet1!A1 — blocked", "⛔ Write blocked — Sheet1!A1 contains 1 non-empty cell(s); nothing was written." A1 still 1. |
| WRITE-04 | pass | "Format Sheet1!A1", "✓ Done"; backup "Format cells · Sheet1!A1". |
| WRITE-05 | pass | See WB-24. |
| WRITE-06 | pass | Header "Inserted 1 row(s) at row 10 in "Sheet1"."; backup "Modify structure · Sheet1!10:10". |
| WRITE-07 | pass | "2 edits" group. |
| WRITE-08 | pass, changed (#789, #791) | See WB-10. The transcript says "Request aborted" instead of "Error: signal is aborted without reason". |
| WRITE-09 | pass | The declined card is red (`data-state="error"`). |
| WRITE-10 | pass, fixed (#794) | Confirm mode: "Write Sheet1!A1 — blocked" with no dialog. |
| WRITE-11 | pass | E10 selected; Pi wrote AJ1:AJ3; the selection was still Sheet1!E10. |
| WRITE-12 | pass, changed (#794) | X1 bold and yellow, X3 italic. Fill `=X1*2` into X2:X5: formulas only, so X2 and X5 were neither bold nor yellow, and X3 kept its italic. `/revert` emptied X2:X5 and left the formats untouched. Same on `a2b3f2d` (R1C1 fill) and `e48d83c` (`copyFrom` paste). |
| WRITE-13 | pass | "Chart 'Test' deleted — no backup", "ℹ️ Backup not created. Excel does not expose the source range…" |
| WRITE-14 | pass | AO8 write held 15 s; a steer sent mid-write showed as a queued "Steer" row; the write completed (AO8 = 1), then Pi answered the steer. |

### workbook/confirm-mode-and-approvals.md

| ID | Result | Note |
| --- | --- | --- |
| APPROVE-01 | pass, changed (#793) | The control now opens a "Workbook changes" popover (Auto / Confirm). Choosing Confirm by real clicks switches at once. There is no toast: the popover replaces the click-to-toggle toast (`e48d83c`). |
| APPROVE-02 | pass, changed (#793) | Tooltip "Auto: Pi applies workbook changes without asking. Click to change." |
| APPROVE-03 | pass, fixed (#793) | With focus left on the trigger, the label read "Confirm" 0.4 s and 1.5 s after the change. |
| APPROVE-04 | pass | Settings "Auto mode" switch (real clicks): "Confirm mode.", then "Auto mode."; the status bar followed. |
| APPROVE-05 | pass | Exactly as tabled: bare `/yolo` and `status` report the mode; `off` twice gives "Confirm mode — …" and changes nothing the second time; `toggle` and `fast` give Auto; `bogus` gives "Usage: /yolo [on\|off\|toggle\|status]". |
| APPROVE-06 | pass, fixed (#793, #794) | Title "Allow this change to the workbook?", note "Confirm mode: Pi asks before each change to the workbook." |
| APPROVE-07 | pass, changed (#794) | The body is "Write Sheet1!D1" with Start / Size / Data; there is no "Tool:" line. |
| APPROVE-08 | pass | Allow once (real click) gave "Edit Sheet1!D1 — 1 changed"; D1 = 5. |
| APPROVE-09 | pass, changed (#794) | See WB-05. |
| APPROVE-10 | pass | A click on the backdrop declined ("The user declined this action in the approval dialog."); 5 backups kept. |
| APPROVE-11 | pass, changed (#794) | Focus goes to the dialog card (`data-autofocus`); the header, app, toasts and command menu are `inert`. See new bug 1 for Tab. |
| APPROVE-12 | pass, changed (#794) | The composer is inert behind the dialog, so `/yolo on` can no longer be typed there (`focus()` refused). Consequence of APPROVE-11. |
| APPROVE-13 | pass, fixed (#794) | D2's dialog showed "1 more waiting", then E2's dialog; both were written. |
| APPROVE-14 | pass, fixed (#794) | Chat 2 H1 and Chat 1 H3: H1's dialog with "1 more waiting", then H3's; H1 = 1 and H3 = 3. H1's dialog had no "from Chat 2" line: the source is computed when the request is made (`src/taskpane/init.ts:1057`), and Chat 2 was still the active tab then. |
| APPROVE-15 | pass | Pi first replied "Please confirm the write of 5 to Sheet1!D1…". |
| APPROVE-16 | pass, changed (#794) | "Allow this change to the workbook?" / "Restore latest backup" / "Action restore". |
| APPROVE-17 | pass, fixed (#794) | Confirm mode: "Let Pi delete backups?" / "Clear backups" / "… Pi asks for this in every mode.", with Delete and Decline. Declining kept every backup. |
| APPROVE-18 | pass, changed (#794) | The dialog's "Stop Pi" closed it and left nothing inert; the card read "Aborted", then "Request aborted"; T5 stayed empty. |
| APPROVE-19 | pass | Auto mode: one dialog, "Allow direct Office.js execution?", with "⚠️ Code references browser capabilities beyond the Excel API: fetch." Escape gave "Office.js execution cancelled by user." (red card). |
| APPROVE-20 | pass | Confirm after CDP `Page.reload`. |

### workbook/backups-and-revert.md

| ID | Result | Note |
| --- | --- | --- |
| REVERT-01 | pass, changed (#794) | "Backups", "Snapshots saved before Pi modifies data", and the notice "Restore any of Pi's changes here. Backups stay on this computer and are kept when you save, up to the limit in Retention settings." |
| REVERT-02 | pass | "No backups yet" / "Pi will save snapshots here before making changes to your data."; the toolbar buttons are hidden. |
| REVERT-03 | pass | "Modify structure · Sheet1!10:10", "just now", "1 change · #e2d19aa2-161". |
| REVERT-04 | pass | "All actions (27)", "Write (20)", "Format cells (1)", "Modify structure (1)", "Restore (5)"; Format cells gives "1 of 27 shown". |
| REVERT-05 | pass | "E2" gives "3 of 27 shown". |
| REVERT-06 | pass | "No backups match the current filter." |
| REVERT-07 | pass | "↓ Newest" becomes "↑ Oldest", oldest first. |
| REVERT-08 | pass, changed (#794) | Restore on the format backup, with no confirmation: A1's bold was removed; toast "Reverted Sheet1!A1 to the backup from just now"; new entry "Restore · Sheet1!A1", "Restored from #b18f32f1-372". |
| REVERT-09 | pass | "Delete this backup?" / "Backup: Sheet1!10:10 (#e2d19aa2-161)"; Escape closed only the confirmation; nothing deleted. |
| REVERT-10 | pass | A confirmed Delete left 26 backups. |
| REVERT-11 | pass | "This will delete 5 backups." gave the toast "Cleared 5 backups" and the empty state. |
| REVERT-12 | pass, changed (#794) | "Reverted Sheet1!R1 to the backup from 7m ago"; R1 empty. |
| REVERT-13 | pass | A second `/revert` put R1 back to 5. |
| REVERT-14 | pass | "No backups for this workbook yet". |
| REVERT-15 | pass, changed (#794) | `/revert` while Pi was streaming ran ("Reverted Sheet1!R1 to the backup from just now"), and the turn kept going. |
| REVERT-16 | pass, changed (#793, #794) | Escape on Backups opened with `/history` closes it; a click on the backdrop closes it too (`e48d83c`). |
| REVERT-17 | pass | Focus returned to the composer. |
| REVERT-18 | pass | "✅ Restored backup {id} (from …) at Sheet1!E2. Changed item(s): 1. Rollback backup created: {id}." E2 empty. |
| REVERT-19 | pass, changed (#794) | Pi wrote L1 = 42 and the user typed 77. `/revert` asked "Overwrite later edits?" ("… 1 cell(s) have changed since Pi's change and would be overwritten: Sheet1!L1 = 77"), with Keep my edits / Restore anyway. Keep left 77; Restore anyway emptied L1. |
| REVERT-20 | pass | 4 gave "Retention limit must be between 5 and 120"; 5 gave "Retention set to 5 backups". |
| REVERT-21 | pass, fixed (#794) | Setting 5 trimmed 44 backups to 5 at once (the newest); after 6 more writes there were still 5. |
| REVERT-22 | pass, fixed (#794) | With 44 stored and the dialog opened fresh: "44 backups", "This will delete 44 backups." See new bug 2 for a dialog left open. |
| REVERT-23 | pass | The list stayed at "25 backups" while Chat 2 wrote AO3 (26 stored); Refresh showed 26, with "Write · Sheet1!AO3" on top. |
| REVERT-24 | pass, fixed (#794) | See WB-12 and WB-13. |
| REVERT-25 | pass | Restore clicked at 6.2 s while another tab's write held the lock until 22.4 s: "Restoring…" until 22.9 s, then "Reverted Sheet1!AM6 to the backup from 1m ago" (`e48d83c`). |
| REVERT-26 | pass, fixed (#794) | T1 = "older", then chart Test created and deleted. The first `/revert` gave "Reverted Sheet1!Test…" and removed the used chart backup; the second reverted T1 (empty); the third put "older" back. |

### workbook/manual-backups.md

| ID | Result | Note |
| --- | --- | --- |
| BACKUP-01 | pass | "Backup command failed: Current workbook identity is unavailable; cannot create or restore manual full-workbook backups safely." (error toast with ×). |
| BACKUP-02 | pass | The same for `/backup restore`. |
| BACKUP-03 | pass | "No manual full-workbook backups for this workbook." |
| BACKUP-04 | pass | `help` and `bogus` give "Usage: /backup [create\|list [limit]\|restore [id]\|clear]". |
| BACKUP-05 | blocked, changed (#794) | `/backup clear` now checks identity before asking, so the confirmation never shows on web-anon. Needs a workbook with an identity (desktop). |
| BACKUP-06 | pass, fixed (#794) | `/backup clear` fails on identity with no confirmation. |
| BACKUP-07 | pass, changed (#794) | `/backup` while Pi was streaming ran (the identity error on web-anon, not "busy"), and the turn kept going. |
| BACKUP-08 | blocked | Desktop only. #796 says the Excel for Mac download fix is still to come. |
| BACKUP-09 | blocked | Desktop only. #794 WT-07 says toasts and lists now show full ids. |
| BACKUP-10 | blocked | Desktop only (WKWebView, WebView2). |
| BACKUP-11 | blocked | Desktop only. |
| BACKUP-12 | blocked | Desktop only. |
| BACKUP-13 | pass | "Download backup" is hidden when there are no backups. |

### Other changes checked at `e48d83c`

- A refused format_cells call stores no backup. Bolding A1 on a protected sheet gave "Error formatting Prot!A1 — no formatting was applied. Sheet "Prot" is protected…", and the backup count stayed at 26.
- Backup-failure note (#794), with the recovery log's `append` made to reject: "⚠️ QuotaExceededError: storage full. This change cannot be reverted from backups.", plus the error toast "Couldn't save a backup. Pi's last change can't be reverted from Backups." The punctuation is now right.

## Regressions and new bugs

No checklist item regressed. Two new bugs were found; both are open on `e48d83c` and were sent to the owning threads.

### 1. Tab leaves the add-in from an overlay (#793; reached at once by #794's approval dialog)

On Excel on the web, Tab from the approval dialog's initial focus moves focus out of the Pi frame into Excel. So does Tab on "Allow once" and Shift+Tab on ×. A keyboard user cannot Tab to Stop Pi, Decline or Allow once. Settings leaks at its ends the same way: on `a2b3f2d`, Shift+Tab from × reached the status-bar mode button behind Settings.

Repro:

1. Confirm mode. Ask Pi to write a cell. The approval dialog opens with focus on the card (`src/ui/approval-dialog.ts:211-212`).
2. Press Tab. `document.hasFocus()` becomes false, and the window blurs.

Cause: `installOverlayFocus` calls `preventDefault()` and moves focus (`src/ui/overlay-dialog.ts:268-289`; 82078ec added the card case). However, Office.js (`excel-web-16.00.js`) has its own keydown handler for keyCode 9 that ignores `defaultPrevented`. It calls `OSF.OUtil.focusToNextTabbable(…, e.target, shift)`. When that finds no neighbour for `e.target` (the card, or the first or last control), it sends `ContextActivationManager_notifyHost(TabExit / TabExitShift)`. Verified fix: a document capture listener that calls `stopPropagation()` when `key === "Tab"` and `defaultPrevented`. With it, Tab cycled × → Stop Pi → Decline → Allow once → ×, and Shift+Tab wrapped from × to Allow once. Sent to #793 (owner) and #794.

### 2. Backups "Clear all" confirms with a stale count (#794, minor)

Repro: open `/history` with 28 backups and leave it open. Have Pi make 14 more writes. Click Clear all. The confirmation says "This will delete 28 backups.", but the toast says "Cleared 42 backups", and all 42 are gone. Opened fresh, the count is right. Cause: `src/commands/builtins/settings-pages/backups-page.ts:228-232` counts `allCheckpoints`, the list loaded when the dialog last loaded, while `onClear()` deletes everything stored. Re-read the count (or reload) before asking. Sent to #794.

## Desktop items for the Mac

WB-16, WB-18, WB-22, BACKUP-05 (needs a workbook with an identity), BACKUP-08, BACKUP-09, BACKUP-10, BACKUP-11, BACKUP-12.
