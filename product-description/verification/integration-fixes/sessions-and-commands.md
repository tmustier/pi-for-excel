# Integration verification: sessions and commands

This file reruns every item in [../sessions-and-commands.md](../sessions-and-commands.md) against the combined fix build, branch `integration-fixes`. That branch is `main` plus PRs #786–#799 (see `INTEGRATION.md` on the branch). The first full pass ran at `a2b3f2d`. After the branch was force-pushed to `e48d83c`, every item whose code changed was rerun there, along with the bugs this pass had reported. The rows give the result at `e48d83c`; where it differs from `a2b3f2d`, the note says so. Items whose code did not change between the two builds were not rerun. After a second force-push to `8099cd6`, the items touched by its changes were rerun again: the tab strip and the tab ×, overlays, Tab and Escape in dialogs, the command menu (now listening on the composer), hints, per-tab drafts and undo, compaction, and the error banner. Rows rerun there say "On `8099cd6`"; every other row is the result at `e48d83c`. All runs were on 27 October 2026, by an agent in an Amp orb.

Expected behavior is the checklist as amended by the fix PRs and `goal.md` "Fixes in flight". Where a PR deliberately changed behavior, the PR wins, and the note names it ("changed (#NNN)").

How it ran:

- **web**: Pi ran in Excel on the web through `node scripts/orb/excel-web-session.mjs --check` (host Excel, platform OfficeOnline, check ok), with the dev server serving this checkout. The taskpane was 349×622 (and 349×662 after a window resize), driven over CDP on its own target: `Input.insertText`, key and mouse events, synthetic `KeyboardEvent`s for Ctrl/Cmd chords the browser would otherwise take, and DOM reads. Reloads used `location.reload()` in the taskpane target, because `Page.reload` works only on top-level targets.
- **web-keys**: real OS keys through `wtype`, on the orb Desktop's Wayland display.
- **Techniques**: model requests were held or failed with CDP `Fetch`. The workbook lock was made observable by wrapping `Excel.run` for `write_cells` with a 8–25 s delay. Compaction used a ~43k-token session (16% of a 272k window) made by having Pi read Sheet1!A10:E909. Workbook identity was stubbed by replacing the global `Office` with an `Object.create` wrapper whose `getFilePropertiesAsync` returned Workbook-A, then Workbook-B. The short pane (340 px) came from resizing the browser window over CDP. Empty-list states came from clearing IndexedDB `sessions`, `sessions-metadata` and `workbook.recentlyClosed.v1.*`.
- **Model turns** ran on GPT-6 Luna (GPT-6.1 Luna is not in the registry), with a few on GPT-6 Sol after a fork. The shared `PI_EXCEL_WEB_AUTH_JSON` login was never logged out or refreshed.

**Result at `8099cd6`: 129 items. 125 pass, 2 fail, 2 blocked (desktop only).** (At `e48d83c`: 127 pass, 0 fail, 2 blocked.)

- Every item that failed on `main` now passes or behaves as its owning PR describes.
- On `a2b3f2d`, two items that passed on `main` failed: SESS-17 (#794) and TABS-22 (#794 + #795). Both are fixed on `e48d83c`. They were not rerun on `8099cd6`, whose diff does not touch the close or lock paths.
- New on `8099cd6`: TABS-05 fails, a regression from #793 @ aff33d6. Clicking the right part of an inactive tab's title closes the tab instead of switching to it.
- MENU-15 fails while a toast is showing: #793's toast stack covers the command menu's lower rows, so a click on one does nothing. This is a #793 bug the earlier passes did not hit, not a change in `8099cd6`.
- New bugs 3, 4 and 5 are fixed on `8099cd6`. All five bugs from the earlier passes are now fixed (see below).

## Results

| ID | Result | Note |
| --- | --- | --- |
| SESS-01 | pass | New workbook upload: one tab "Chat 1", empty state, focus in the composer. |
| SESS-02 | pass | Chat 1 on GPT-6 Luna (not the default Sol); "+" gives Chat 2 on GPT-6 Luna, "High". The inherited model differs from the default, so this time the inheritance is checked. |
| SESS-03 | pass, changed (#789) | Drafts are per tab: `/new` typed in Chat 1; Chat 2's composer is empty; Chat 1 still holds `/new`. The suspected bug is fixed. On `8099cd6` each tab gets its own textarea, and per-tab drafts still survive switching. |
| SESS-04 | pass, changed (#789) | Drafts survive a reload: a hint's text came back in Chat 2 after the reload, and Chat 1 still had `/new`. |
| SESS-05 | pass | Chat 4 (the last tab) closed; "+" gives Chat 5. |
| SESS-06 | pass, changed (#795) | "Q3 budget" and "Chat 7" are still "Q3 budget" and "Chat 7" after a reload: numbers and chosen titles persist. |
| SESS-07 | pass | Same order, Chat 7 active, Q3 budget transcript intact. |
| SESS-08 | pass | Blank Chat 7 shows the empty state after the reload. |
| SESS-09 | pass, changed (#795) | Resuming the open session focuses its tab: "Already open: Q3 budget", no second tab. The suspected bug is fixed. |
| SESS-10 | pass, changed (#795) | Sessions save from the first user message, so a closed blank tab is listed only under Recently closed ("Chat 29 · Closed just now · Reopens in new tab"). Saved sessions shows "No available sessions for this workbook." |
| SESS-11 | pass | Busy dot on the active and the background tab, gone at idle. A background tab that finished then shows #793's "Finished, not viewed yet" marker. |
| SESS-12 | pass | `/new` while Chat 7 is busy gives Chat 8; Chat 7 keeps its dot and finishes. |
| SESS-13 | pass, changed (#789) | The error banner is now per tab. A turn failed in Chat 8 (`Fetch.failRequest`) shows the banner there; it disappears on Chat 7 and returns on Chat 8. Same on `8099cd6`. The banner still shows the CORS-proxy text: #796's "the browser blocks the proxy here" wording shows only when `isProxyBlockedByBrowser()` is true (a loopback proxy URL is configured and blocked by frame policy), which was not the case in this run. |
| SESS-14 | pass | As corrected: `document.url` and `getFilePropertiesAsync` both return "", and there is no filter row. |
| SESS-15 | pass, changed (#789) | Background `/compact` with the request held: "Compacting context…" shows only on the compacting tab, while the other tab's indicator is empty. Rerun on `e48d83c` with the same result. |
| SESS-16 | pass, changed (#795) | Fork on: an empty Chat 9 switched to gpt-6-sol in place. A non-empty Chat 8 still forks: "Opened Chat 8 (gpt-6-sol) in a new tab". |
| SESS-17 | pass on `e48d83c` (fail on `a2b3f2d`) | On `a2b3f2d` the waiting tab never showed that it was waiting, during a 25 s overlap: a regression from #794's coordinator (see below). On `e48d83c` (#794 @ a142d28) the waiting tab shows the badge "Waiting" and the status bar "Waiting for workbook" (#793 renamed "Locking…"). The holder has its × disabled ("Wait for workbook changes to complete") and "Changing workbook". The writes ran in turn. |
| SESS-18 | pass, changed (#793) | The busy dot is visible on an inactive tab. |
| SESS-19 | pass | Escape from the idle composer moves focus to the active tab button. |
| SESS-20 | pass | Name "Open tab Chat 8", tooltip "Chat 8". #793 appends the state to the name (", Working", ", Finished, not viewed yet"). |
| SESS-21 | pass | Escape in Chat 8 shows "Request aborted"; Chat 7 stays busy. |
| SESS-22 | pass, changed (#795) | A workbook uploaded after a Chrome restart opened with only Chat 1, although the previous workbook had 6 tabs. Layouts are keyed by document token (`workbook.tabLayout.v1.doc_instance:…`). |
| TABS-01 | pass | "+" adds Chat 13 at the right end, active, with the empty state and focus in the composer. |
| TABS-02 | pass | `/new` while busy opens Chat 16, with no refusal. |
| TABS-03 | pass | web-keys (wtype) on `e48d83c`: Ctrl+T from the composer opened `chrome://newtab`; Pi's tabs were unchanged. |
| TABS-04 | pass, changed (#795, #793) | Tab shortcuts are off while a dialog is open: Ctrl+T delivered to the page with `/resume` open adds no tab, and the picker stays. |
| TABS-05 | fail on `8099cd6` (pass on `e48d83c`), regression (#793 aff33d6) | Clicking the left part of an inactive tab's title still switches to it. Clicking the right part closes the tab: with "Chat 12" inactive (title 117–158 px), the hover × sits at 141–163 over the title, and a click at 152 gave "Closed Chat 12 · Undo" instead of a switch. See regression 3. |
| TABS-06 | pass | ArrowRight from the last tab wraps to the first; ArrowLeft goes back; focus stays on the tab. Rerun on `e48d83c` with 11 tabs. Plain arrows inside the tab menu no longer switch tabs (#793). |
| TABS-07 | pass | Ctrl+PageUp in the composer does nothing; Ctrl+PageDown from the body switches. Cmd+Shift+[ does nothing in the composer and switches from the body; ] switches back. |
| TABS-08 | pass, changed (#795) | With focus on a tab, Ctrl+PageUp and Ctrl+PageDown switch tabs. The suspected bug is fixed. |
| TABS-09 | pass | web-keys (wtype): two Pi tabs, focus on the body, Ctrl+PageDown leaves the active tab unchanged. Typing "xy" with wtype confirmed the keys reach the page. |
| TABS-10 | pass | Right-click on the first tab while another is active: six items, "Move left" disabled, active tab unchanged. The menu takes focus (#793). |
| TABS-11 | pass, changed (#789, #793) | Escape with the tab menu open while responding closes only the menu, and the turn continues; a second Escape stops it ("Request aborted"). The checklist expected the first Escape to stop the turn; the PRs made the topmost overlay take Escape. |
| TABS-12 | pass | Escape while idle closes the menu. |
| TABS-13 | pass | "Chat 13 copy" added at the right end, active; toast "Duplicated Chat 13". |
| TABS-14 | pass | "Wait for this tab to finish before duplicating it" (wording now ends in "it"); no new tab. |
| TABS-15 | pass | Order changed; toast "Moved tab left". |
| TABS-16 | pass | With one tab: no ×; Rename and Duplicate enabled; moves, Close other tabs and Close tab disabled; Ctrl+W gives "Can't close the last tab". |
| TABS-17 | pass | "Stop and close this tab?" / "Pi is still responding in this tab. Stop and close it?" / Keep open / Stop and close, from × and from Ctrl+W. Rerun on `e48d83c`. |
| TABS-18 | pass | Escape and "Keep open" cancel the close, and the tab stays busy. Rerun on `e48d83c`. |
| TABS-19 | pass | Stop and close on the first tab: the neighbour becomes active, toast "Closed Chat 15 · Undo". On `a2b3f2d` it also showed a misleading "Stopping — closing after Pi's current change finishes." with no write running. Fixed on `e48d83c` (#794): only "Closed Chat 30 · Undo". |
| TABS-20 | pass | "Closed 3 other tabs" (and later 5, 8 and 10). |
| TABS-21 | pass | While a tab holds the lock: × disabled ("Wait for workbook changes to complete"); Ctrl+W gives "Wait for workbook changes to finish before closing this tab"; "Close tab" disabled. |
| TABS-22 | pass on `e48d83c` (fail on `a2b3f2d`), changed (#794) | On `a2b3f2d` a confirmed "Stop and close", after the tab took the lock behind the dialog, showed "Wait for workbook changes to finish…", did not stop the turn, and left the tab open. This was the #795 + #794 combination; see below. On `e48d83c` the toast is "Stopping — closing after Pi's current change finishes." and the tab closed 0.8 s later. Neither the delayed L1 write nor the queued M1 write landed: my delay sat before the call reached Excel, so Stop cancelled it there. A write that is already inside Excel and then lands was not observed. |
| TABS-23 | pass, changed (#789) | Closing a compacting tab now asks "Stop and close this tab?". The suspected bug is fixed. Rerun on `e48d83c`. |
| TABS-24 | pass, changed (#795, #793) | Ctrl+W (and Ctrl+Shift+T) delivered to the page with the picker open are ignored: no tab closes, and the picker stays. |
| TABS-25 | pass | web-keys (wtype): with Excel the only page, Ctrl+W closed the Excel browser tab (only `about:blank` was left); Pi saw nothing. |
| TABS-26 | pass, changed (#793) | On `8099cd6`: with 11 tabs at 349 px, tabs keep their natural width and full titles (no more "C…"), and the strip scrolls with ‹ ›. The native scrollbar is gone (0 px). The arrows scroll the strip and it stays where it was scrolled. "+" stays outside the scroller and is always visible. ArrowRight wraps and scrolls the active tab into view. On `a2b3f2d` the new tab was half under › and "+" was off-screen; fixed. |
| TABS-27 | pass | After the duplicate, a new tab is "Chat 15". |
| TABS-28 | pass, changed (#789) | Compaction held, × → Stop and close: "Compaction stopped", "Closed Chat 32 · Undo"; `/reopen` gives "Reopened: Chat 32", still uncompacted, as #789 intends (Stop cancels compaction). A compaction that finishes is saved at once. On `8099cd6`: Stop and close gives "Compaction stopped" in place of the "Compacting…" toast (new bug 5 fixed), and the reopened tab is uncompacted at 11%. `/compact` takes 11% → 1%, still 1% after a reload. On `a2b3f2d` the compacted history was not saved (see the fixed combination bug below). |
| RENAME-01 | pass | Label "Q3 budget", name "Open tab Q3 budget". The toast is now "Renamed tab to “Q3 budget”" (#793, one wording per event). |
| RENAME-02 | pass | `/name Foo` while busy: "Wait for this tab to finish before running /name" (new wording, replacing "Can't run /name while Pi is busy"). The text is kept. |
| RENAME-03 | pass, changed (#793) | Double-click opens the in-pane "Rename tab" dialog with the field focused and prefilled; Enter gives "Renamed tab to “Revenue Model”". The suspected bug is fixed. |
| RENAME-04 | pass, changed (#793) | "Rename tab…" opens the same dialog. Empty text gives "Tab name reset to automatic", and the label returns to "Chat 18". |
| RENAME-05 | pass | From the menu and from a double-click: "Wait for this tab to finish before renaming it". |
| RENAME-06 | pass | `/name` from the menu: "Usage: /name My Session Name"; the composer is empty. |
| RENAME-07 | pass, changed (#795) | After a reload the label is still "Q3 budget". The suspected bug is fixed. |
| RENAME-08 | pass, changed (#795) | Resuming the closed "Q3 budget" opens a tab labelled "Q3 budget"; toast "Resumed in new tab: Q3 budget". |
| RENAME-09 | blocked | desktop. The `window.prompt` freeze on Excel for Mac should be gone, since #793 uses an in-pane dialog; this needs a Mac run. |
| REOPEN-01 | pass, changed (#795) | Undo reopens with the old number: "Chat 16", toast "Reopened: Chat 16". The draft in Chat 19 is untouched. |
| REOPEN-02 | pass | Chat 18 closed; "draft text" typed in Chat 19; Ctrl+Z gives "Reopened: Chat 18". Chat 19's draft is still "draft text". |
| REOPEN-03 | pass | Ctrl+Shift+T delivered to the page from the composer: "Reopened: Chat 15". |
| REOPEN-04 | pass | web-keys (wtype): with Pi's Chat 2 closed and a browser tab closed, Ctrl+Shift+T reopened the browser tab; no Pi tab. |
| REOPEN-05 | pass, changed (#795, #793) | Close, then Ctrl+Shift+T at once: the Undo toast is retired and "Reopened: Chat 18" shows. The suspected bug is fixed. |
| REOPEN-06 | pass, changed | With the stale Undo toast gone (REOPEN-05), there is nothing left to click. |
| REOPEN-07 | pass | With the list cleared, `/reopen` and Ctrl+Shift+T give "No recently closed tab". |
| REOPEN-08 | pass | With the list empty, "abc" then Ctrl+Z gives "" and no tab. |
| REOPEN-09 | pass | With the toast gone, "abc def" then Ctrl+Z gives "abc" and no tab. On `8099cd6`, Ctrl+Z in another tab no longer brings back this tab's text (new bug 4 fixed). |
| REOPEN-10 | pass, changed (#795 CL-02) | Ctrl+Z reopens only while the Undo toast is showing. With the toast gone and an empty composer, nothing reopens. |
| REOPEN-11 | pass | Newest first, "Closed just now · Reopens in new tab". With "Replace current tab" selected, a real click on Chat 15 opened a new tab: "Reopened: Chat 15". |
| REOPEN-12 | pass | After Close other tabs, the rightmost closed tab comes back first. |
| REOPEN-13 | pass, changed (#795) | Recently closed is now persisted per workbook: after a reload, `/reopen` gives "Reopened: Chat 16" with its old label. |
| RESUME-01 | pass | Title, subtitle, "Open in new tab" (selected) / "Replace current tab", hint, Recently closed and Saved sessions. |
| RESUME-02 | pass | Gear → Resume session opens the same picker. |
| RESUME-03 | pass | No workbook filter (identity unknown). |
| RESUME-04 | pass, changed (#793) | The picker takes focus on its card, so typing does not reach the composer; the first Escape closes it. On `8099cd6`, Tab and Shift+Tab wrap inside the picker, Tab from the last item goes to ×, and focus stays in the add-in (new bug 3 fixed). Escape in a dialog's text field (Rename tab, Copy text) first moves focus to the card, and a second Escape closes the dialog. |
| RESUME-05 | pass, changed (#793) | With focus in the picker, Enter sends nothing to the tab behind it. |
| RESUME-06 | pass | "Resumed in new tab: Q3 budget"; the tab keeps the session's label (#795). |
| RESUME-07 | pass, changed (#795) | As SESS-09: the session's existing tab is focused, with no second tab. |
| RESUME-08 | pass, changed (#795) | "Replace current tab" on a busy tab: "Wait for this tab to finish, or open the session in a new tab". Nothing is replaced, and the picker stays open. |
| RESUME-09 | pass | `/resume-here x` while busy: "Wait for this tab to finish before running /resume-here"; the text is kept. |
| RESUME-10 | pass | `/resume` opens the picker while busy. |
| RESUME-11 | pass | Idle "Replace current tab": "Resumed in current tab: Chat 8 (gpt-6-sol)". The transcript and model (GPT-6 Sol) swap, and the label takes the session's title (#795). |
| RESUME-12 | pass | Identity stubbed (Workbook-A, then Workbook-B): the filter row "Show sessions from all workbooks / Workbook-B.xlsx"; A3-session is hidden until the box is ticked. Choosing it gives "Resume session from another workbook?"; Escape returns to the picker; "Resume anyway" gives "Already open: A3-session" (#795). |
| RESUME-13 | pass | Responding, picker opened from the gear, Escape: the picker closes and the turn runs to "OK". |
| RESUME-14 | pass | With the sessions stores and recently closed cleared: "No previous sessions", and no picker. |
| RESUME-15 | pass | "just now", "13m ago", "Closed 11m ago". |
| RESUME-16 | pass on `e48d83c` (fail on `a2b3f2d`) | On `a2b3f2d`, with several recently closed tabs, the saved list was 0 px tall at 349×622/662 and could not be clicked (new bug 1). On `e48d83c` it is 414 px, and a real pointer click on a saved entry resumes it ("Already open: Chat 32"). |
| RESUME-17 | pass, changed (#789, #793) | `/resume` while responding, then Escape: the first Escape closes only the picker and the turn continues; a second Escape stops it. The suspected bug is fixed (the topmost overlay takes Escape). Rerun on `8099cd6`. This row was missing from the earlier commit of this file. |
| RESUME-18 | pass, changed (#793) | At 340 px tall the card scrolls (1719/270 px) and the saved list is full height. |
| EXPORT-01 | pass, changed (#794 47a96eb) | `/share-session x` while busy is refused and keeps its text. `/copy` and `/export` now run while busy ("Transcript download started (11 messages)"). |
| EXPORT-02 | pass | "Session sharing coming soon". |
| EXPORT-03 | pass | After a stopped response, `/copy` gives "No assistant message to copy". |
| EXPORT-04 | pass, changed (#795) | `/copy` gives "Copied to clipboard" through the `execCommand("copy")` fallback (copied "OK"). |
| EXPORT-05 | pass, changed (#795) | With `execCommand` also failing, the manual "Copy text" dialog opens with the text selected. The suspected bug is fixed. |
| EXPORT-06 | pass, changed (#795, #796) | "Transcript copied (2 messages, 90KB)". With no clipboard: "The clipboard isn't available here, so the transcript download started instead (2 messages)", and the download completed. |
| EXPORT-07 | pass | `/export audit clipboard`: "Audit log copied (9 records, 4KB)". |
| EXPORT-08 | pass, changed (#795) | "Transcript download started (2 messages)"; `pi-session-2026-10-27.json` completed (91,752 B); no extra browser tab. |
| EXPORT-09 | pass | "Audit log download started (9 records)"; `pi-audit-log-2026-10-27.json` (4,545 B). |
| EXPORT-10 | pass, changed (#795) | Counts now cover only visible messages: the toast says 2, while the file's `messageCount` is 5 (system, 2× autoContext, user, assistant). |
| EXPORT-11 | blocked | desktop. #796 replaces the Mac download path (`saveBlobAsFile`); Excel for Mac needs to be clicked through. |
| MENU-01 | pass, changed (#795) | Order as documented, without `/default-models` (removed); `/snake` [ext] last. |
| MENU-02 | pass | `/tab` lists `/new`, `/resume-here`, `/reopen`. |
| MENU-03 | pass | `/skills` and `/plugins`: the exact match is first. |
| MENU-04 | pass | Escape closes the menu and keeps the text; typing reopens it. Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-05 | pass | Tab fills without running: `/resu` → `/resume `. (`/res` highlights `/experimental` first because "features" matches; same as `main`.) Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-06 | pass, changed (#789) | `/foo` + Enter keeps the draft and shows "No command /foo. Start with a space to send it as a message". Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-07 | pass, changed (#789) | A chosen command keeps its arguments: `/name the current` + Enter with the menu open renames the tab to "the current". Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-08 | pass | Typed and refused while busy: the text is kept. |
| MENU-09 | pass, changed (#789) | Chosen from the menu while busy: the refusal toast shows, and the composer keeps `/name`. The suspected bug is fixed. Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-10 | pass | On a blank tab: "Too few messages to compact". |
| MENU-11 | pass | `/new` + Enter with the menu open while responding opens Chat 19; nothing is steered. Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-12 | pass, changed (#789) | `/` then Escape while responding: the menu closes, and the turn continues. Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-13 | pass, changed (#789, #793) | `/new` in the menu, then a click on another tab: that tab's own (empty) draft shows, the menu is gone, and Enter does nothing. Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-14 | pass, changed (#793) | `/new` in the menu, then gear → Resume session: focus is in the picker, and Enter runs nothing. Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-15 | fail on `8099cd6` while a toast shows (#793) | With no toast showing, a click on the `/share-session` row runs it ("Session sharing coming soon"). Right after another command's toast, the toast stack sits over the menu's lower rows (menu y 532–576), so the click hits the toast and nothing runs. It works again once the toasts expire. See new bug 6. |
| MENU-16 | pass | `/mod` + Shift+Tab gives `/model `, and thinking is unchanged. Rerun on `e48d83c` and `8099cd6`. |
| MENU-17 | pass, changed (#789) | The menu closes on an outside click (blur) and when the gear menu opens. Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-18 | pass | Hover moves the highlight; 0 DOM mutations in 1 s (no continuous re-render). Rerun on `8099cd6` (menu now listens on the composer): same. |
| MENU-19 | pass, changed (#795) | `/default-models` and its Ctrl+P hint are removed. |
| HINTS-01 | pass | The tagline and the four hints, in order. |
| HINTS-02 | pass, changed (#793) | No tooltip: the hints are `aria-describedby` a visible note, "A hint fills the message box so you can edit it before sending." |
| HINTS-03 | pass | Previews are 118 characters. |
| HINTS-04 | pass, changed (#789) | The hint is appended to "my draft" after a blank line; caret at the end, focus in the composer, nothing sent. Rerun on `8099cd6` (menu now listens on the composer): same. |
| HINTS-05 | pass, changed (#789) | Ctrl+Z removes the hint, leaving "my draft". Rerun on `8099cd6` (menu now listens on the composer): same. |
| HINTS-06 | pass, changed (#789) | After a reload, the hint's text is restored in the composer, and the empty state shows. HINTS-04 to 06 rerun on `8099cd6` with the same result. |
| HINTS-07 | pass | At 340 px: the π mark and the previews are hidden (`display: none`); the labels show. |
| HINTS-08 | pass | As corrected: at 340 px the gear menu covers the hints, and a click at a hint's position opened Files. |
| HINTS-09 | pass | With a provider connected, the empty state was gone 37 ms after Enter (request held). The no-provider case was not run. |

## Regressions and new bugs

Each was sent to its owning PR thread when found. Status is at `8099cd6`.

### Regressions (passed on `main`)

1. **SESS-17: a waiting tab never showed that it was waiting.** Fixed in #794 @ a142d28.
   - Repro: delay `write_cells`' `Excel.run` by 25 s. Start a write in tab A, then one in tab B. While one holds the lock, the other shows only its busy dot: no badge on the tab and nothing in the status bar.
   - Cause: `src/workbook/coordinator.ts:229` emitted `queued` before `state.queue.push` (:240). The new listener in `src/taskpane/init.ts:1904-1916` recomputes every tab from `getSnapshot()`, which did not yet contain the queued write, so the waiting session came out `idle`.
   - Owner: #794.
2. **TABS-22: "Stop and close" neither stopped nor closed once the tab held the lock** (only in the combination). Fixed in #795 @ e174347.
   - Repro: delay `write_cells` 15 s. Ask for a one-cell write, click × at once, and confirm when "Changing workbook" shows. Result: toast "Wait for workbook changes to finish before closing this tab", the turn runs to "OK", the write lands, and the tab stays open.
   - Cause: #795's post-dialog lock recheck (`src/taskpane/init.ts:1717`) returned before #789's `actionQueue.stop()` and #794's wait-then-close branch (:1729).
   - Owners: #795 + #794; integration was told.
3. **Open: TABS-05: clicking an inactive tab's title can close it instead of switching to it.** New on `8099cd6` (#793 @ aff33d6).
   - Repro: two tabs, say "menutest" (active) and "Chat 12". Move the mouse over the right half of "Chat 12"'s title and click. The tab closes ("Closed Chat 12 · Undo"). Clicking the left part of the title switches as before.
   - Measured: tab 108–167 px, title 117–158, × 141–163 once hovered; a click at 152 closed it.
   - Cause: `src/ui/theme/components/tabs.css:255-265`. An inactive tab's × is now `position: absolute; right: 0` over the title, hidden with `opacity: 0; pointer-events: none`. On hover it gets `pointer-events: auto`, so the first click a mouse user makes on that part of the title lands on ×. Before, the × took its own space.
   - Owner: #793.

### New bugs

1. **The resume picker's Saved sessions list collapsed to 0 px** in a normal-height pane once a few tabs had been recently closed (RESUME-16). Fixed in #793 @ 82078ec.
   - Repro: 349×622 or 349×662, five or six recently closed tabs, then `/resume`. The saved list is 0 px tall, and a click there lands on the backdrop and closes the picker.
   - Cause: `src/ui/theme/overlays/provider-resume-shortcuts.css:20-23` (`flex: 1 1 auto; min-height: 0`) inside the 80vh card.
   - Owner: #793.
2. **Compaction was not saved when it archived every user message** (only in the combination; TABS-28, #789's "compaction is saved at once"). Fixed in #789 @ 79a4136 and #795 @ 6e1ba25.
   - Repro: a one-turn session with a big read, at 16%. `/compact` takes it to 2%, but the IndexedDB record is unchanged and a reload brings back 16%.
   - Cause: #789's `runCompact` saved without force (`src/taskpane/init.ts:1316-1324`). #795's gate in `src/taskpane/sessions.ts:192` skipped the save when no top-level user message remained.
   - Owners: #789 + #795; integration was told.
3. **Tab moved focus out of the add-in from an open overlay.** Fixed on `8099cd6` (#793 @ aff33d6): Tab and Shift+Tab wrap inside `/resume`, `document.hasFocus()` stays true, and Tab from the last item reaches ×. Tab in the gear menu returns to its trigger.
   - Repro: in Excel on the web, open `/resume` (focus is on the card) and press Tab once. `document.hasFocus()` becomes false, and focus goes to the Excel host. The same happens with Tab from the last control. Shift+Tab from × lands on the card instead of wrapping.
   - History: #793 @ 82078ec added the `active === card` case (`src/ui/overlay-dialog.ts:284`), but the bug still reproduced on `e48d83c`.
   - Cause: Office.js adds a window keydown (bubble) listener (OSF). For Tab it ignores `defaultPrevented` and calls `OSF.OUtil.focusToNextTabbable(…, e.target, …)`, or `ContextActivationManager_notifyHost(TabExit)`, based on the element focused before the trap ran. That overrides the trap's `first.focus()`. The trap in `src/ui/overlay-dialog.ts:269-289` needs `event.stopPropagation()` and must move focus itself on every Tab.
   - Owner: #793.
4. **Ctrl/Cmd+Z in one tab's composer brought back text typed in another tab.** Fixed on `8099cd6`: each tab gets a fresh textarea when its draft loads. Ctrl+Z in tab Y no longer gives X's text, Y's own undo still works, and per-tab drafts survive switching and a reload.
   - Repro on `e48d83c`: type "alpha beta" in tab X, clear it, and switch to an empty tab Y. Ctrl+Z gives "alpha beta"; a second Ctrl+Z gives "alpha". Y keeps that text as its own saved draft.
   - Cause: every tab shares one textarea, and `src/taskpane/runtime-sidebar-binding.ts:25-26` swaps drafts by setting `input.value` (`src/ui/pi-input.ts:51-58`). Chrome keeps the textarea's native undo history across the swap. #789's hint insertion deliberately relies on that history (`src/ui/pi-sidebar.ts:1358-1363`).
   - Owner: #789.
5. **"Compacting to free up context" stayed on screen after compaction ended.** Fixed on `8099cd6` (integration fix): "Compaction stopped" now replaces the progress toast, also when Stop and close stops it.
   - Repro on `a2b3f2d`: with #793's stacked toasts, the 60 s progress toast stayed after "Summarized…".
   - History: #793 @ 82078ec gave the compaction toasts an id, so success, "Nothing to compact" and failure replaced it. Still open on `e48d83c`: a stopped compaction shows "Compaction stopped" without the id (`src/commands/builtins/export.ts:851`), so the progress toast stays up to 60 s. It was still showing 10 s after the stop, across closing and reopening the tab.
   - Owner: #793.

6. **Open: toasts cover the slash-command menu, so a click on a covered row does nothing** (MENU-15).
   - Repro: run a command that shows a toast (for example `/share-session`), then at once type `/` and click a row in the lower part of the menu. Menu y 532–576 at 349×662; the toast stack covers those rows, so the click hits the toast and no command runs. It works once the toasts expire.
   - Cause: `src/ui/theme/components/toasts.css:11-12` places the stack at `bottom: calc(var(--pi-dock-height) + 8px)`, just above the dock where the command menu opens, at `z-index: var(--z-toast)`, which is above the menu's `var(--z-popover)` (`src/ui/theme/components/menus.css:80`). Unchanged since #793's b25a6173; earlier passes clicked with no toast showing.
   - Owner: #793.

Minor (sent to #793; fixed on `8099cd6`): at 11 tabs, every 60 px tab read "C…", and a native horizontal scrollbar showed under the tab strip.

## Desktop items (for the Mac)

- **RENAME-09**: double-click and "Rename tab…" in Excel for Mac now open #793's in-pane dialog instead of `window.prompt`. Check that it no longer freezes, and the "Renamed tab to …" and "Tab name reset to automatic" toasts.
- **EXPORT-11**: `/copy`, `/export` and `/export download` in Excel desktop. #795 adds the `execCommand` and manual-dialog copy fallbacks; #796 makes `saveBlobAsFile` (blob `<a download>`, no `window.open`) the only download path. Check that a file is actually saved.
- **Desktop counterparts of web items**: Cmd+T, Cmd+W and Cmd+Shift+T should now be ignored while a dialog is open (TABS-04, TABS-24); check that Excel does not take them instead. Check the Tab trap in overlays on WKWebView (new bug 3 was Office-on-the-web specific and is fixed there on `8099cd6`), and that clicking an inactive tab's title switches to it (regression 3). REOPEN-04 desktop: whether the stale "Closed …" toast is now retired (#795).
