# Integration verification: ui and cross-cutting

This file reruns every item in [../ui-and-cross-cutting.md](../ui-and-cross-cutting.md) against the combined fix build. The build is branch `integration-fixes`, which is `main` plus PRs #786–#799 (see `INTEGRATION.md` on that branch). An agent in an Amp orb ran it on 27 October 2026. The pass started on `a2b3f2d`. The branch was then force-pushed twice, and the changed items were rerun on `e48d83c` and `8099cd6`. The results below are for `8099cd6` unless a note says otherwise. Expected behavior is the checklist as amended by the fix PRs. Where a PR deliberately changed behavior, the PR wins, and the note names the PR.

How it ran:

- **web**: Pi ran in Excel on the web through `node scripts/orb/excel-web-session.mjs --check`, with the dev server started from the `integration-fixes` checkout (host Excel, platform OfficeOnline, write and read-back ok). The taskpane was 349 px wide (342 px, then 580 px tall). It was driven over CDP on its own target with `Input.insertText`, key and mouse events, and DOM and IndexedDB reads.
- **web-keys**: real OS keys sent with `wtype` into the headed Excel Chrome on the orb Desktop.
- **Reloads**: reloads used `location.reload()` in the taskpane target, because CDP `Page.reload` is refused on an out-of-process iframe ("Command can only be executed on top-level targets").
- **web-2**: a second browser tab on the same workbook link in the same profile. Opening it crashed pane A's renderer once (sad-face pane, probably the Excel Chrome's memory cap), and pane A's frame was reloaded.
- **Injected states**: `Excel.run` was wrapped in the taskpane to delay calls (slow write). Provider failures (503, 400 context overflow, 400 "cancelled", 5 s delays) were substituted with CDP `Fetch`. Offline used `Network.emulateNetworkConditions`; a second CDP session resets `navigator.onLine`, so one session was used. Storage failures were made by throwing from `IDBObjectStore.prototype.put` for one key or store. The unsaved-workbook case removed `pi.workbookInstanceId` from the document settings, which were restored afterwards. A wrong OpenAI (API) key was written straight into the `provider-keys` store, because #795 now rejects a 401 key at Save.
- **Fresh browser contexts**: CDP `Target.createBrowserContext` was used for the welcome screen (`/__pi-auth` blocked) and for empty storage.
- **Model turns**: real turns on the shared ChatGPT login, mostly on GPT-6 Luna (GPT-6.1 Luna is not in the registry), plus a few on GPT-6.1 Sol. The shared credential was never logged out or refreshed.
- **Screen reader**: no screen reader was available. Live regions were checked by recording `[aria-live]` regions and `role=alert`/`role=status` nodes with a MutationObserver, and by checking that the regions are not inert.

**Result on `8099cd6`: 118 items, 104 pass, 0 fail, 14 blocked.** Thirteen blocked items need Excel desktop (listed at the end). MULTI-10 is blocked to protect the shared login. Every item that failed on `main` now passes as its owning PR describes. Four checks that passed on `main` failed during the pass and are fixed on `8099cd6`: STATUS-14, STATUS-15, PROVIDER-11, and the focus part of UTIL-02. Three minor new bugs are still open (see below).

## Results

| ID | Result | Note |
| --- | --- | --- |
| STATUS-01 | pass (#793) | The bar shows "π GPT-6.1 Sol", "High", "1%", "Auto". At 349 px the "/ 500k" limit is hidden by #793's one-row rule (`@container (max-width: 360px)`); the context popover still shows "3,494 / 500,000 tokens". |
| STATUS-02 | pass (#793) | Model, thinking and context tooltips are as quoted. The mode tooltip is now "Auto: Pi applies workbook changes without asking. Click to change." (the mode control is a popover, #793). |
| STATUS-03 | pass (#795) | While busy, the picker opens read-only: 7/7 rows are `aria-disabled`, with the note "Pi is working in this tab. Change the model when it finishes, or press Stop." A click leaves the model unchanged and shows no toast. New minor bug: rows stay disabled after the turn ends if the picker is still open (see below). |
| STATUS-04 | pass (#795) | GPT-6 Luna lists Off…Max with ✓ on High. GPT-4.1 mini (OpenAI (API)) lists only "Off ✓". Switching that tab from High clamps the level to Off (#795). |
| STATUS-05 | pass | A real click on "Low" closes the popover, the label shows "Low", there is no toast, and focus returns to the trigger. |
| STATUS-06 | pass | web-keys: Shift+Tab ×4 in the composer goes High → Extra high → Max → Min → Low, with the toast "Thinking: {level} (next turn)" each time. Focus stays in the composer. |
| STATUS-07 | pass | web-keys: during a streaming turn, Escape closes only the thinking popover. Stop stays and the turn keeps streaming. |
| STATUS-08 | pass (#793) | The suspected bug no longer reproduces: switching tabs with a popover open (Ctrl+PageDown) closes the popover, so it cannot change the hidden tab. On `a2b3f2d`, ← or → inside the popover switched Pi tabs (new bug, fixed in `e48d83c`). |
| STATUS-09 | pass | "Context usage", the description, "3,494 / 500,000 tokens", "/compact" and "/new". Focus moves into the popover (#793). |
| STATUS-10 | pass | By keyboard: ↓ then Enter on "/new" opens "Chat 3", with focus in the composer. |
| STATUS-11 | pass (#789) | At 11%: "Nothing to compact". At 34%: "Compacting to free up context", then "Summarized 2 messages", and the meter redraws at once (34% → 18%, #789). |
| STATUS-12 | pass (#789) | GPT-6 Luna (272k), filler messages: 31% and 50% grey; 70% yellow with "Context 70% full." (#789's per-model budget). The next send auto-compacted before sending ("Summarized 8 messages", 34%), so red and over-100% were not reached. |
| STATUS-13 | pass (#793) | During a write the badge reads "Changing workbook", with the tooltip "This tab is changing the workbook. Other tabs wait until it finishes.", and disappears after (#793 wording). |
| STATUS-14 | pass on `e48d83c` | On `a2b3f2d` the waiting tab never showed it was waiting (regression, #794, see below). On `e48d83c`, the background waiting tab reads "Chat 3 Waiting" (aria "Working, Waiting") while the active tab shows "Changing workbook". |
| STATUS-15 | pass on `e48d83c` | Same fix. With the waiting tab active, the bar shows "Waiting for workbook" with the tooltip "A workbook change is queued behind another tab.", then "Changing workbook" once the other write ends. |
| STATUS-16 | pass (#793) | The mode control opens a popover ("Workbook changes": Auto / Confirm, each with a line). Keyboard ↓, Enter sets "Confirm" at once with no toast, focus goes back to the trigger, and `execution.mode.v1` is "safe". |
| STATUS-17 | pass (#793) | Behavior changed: the bar now renders at once. With the pointer resting on the bar (`:hover` true), "Changing workbook" went away as soon as the write ended. |
| STATUS-18 | pass (#793) | Changed: in a 349 × 580 pane the thinking popover opens at y 50–538 and no longer covers the meter (`elementFromPoint` on the meter hits the meter). Clicking the meter swaps to the context popover. |
| STATUS-19 | pass (#795) | With ChatGPT and OpenAI (API) both connected, labels read "GPT-6 Luna · ChatGPT" and "GPT-4.1 mini · API". The suspected bug is fixed. |
| STATUS-20 | pass | "No active session" was never seen. Sampling every 25 ms during a reload, the bar was empty until 0.9 s and then "π GPT-6 Luna High 5% Auto" (as the corrected document says). |
| STATUS-21 | blocked | Needs Excel desktop. |
| RENDER-01 | pass | `<b>raw</b> and $5 and $6` shows as text in the user message and the answer: no `<b>`, no KaTeX. |
| RENDER-02 | pass | Link "image: logo" to the image URL; no `<img>`. |
| RENDER-03 | pass | "js" is plain text; "site" has `target=_blank`, `rel="noopener noreferrer"`. |
| RENDER-04 | pass (#786) | The suspected bug is fixed: "line one" and "line two" are on separate lines (5 `<br>` in the user message, `breaks: true`). |
| RENDER-05 | pass | The table sits in `.pi-md-table-wrap` with `overflow-x: auto`. |
| RENDER-06 | pass | The `python` fence gets a bar with "python", "Copy" and highlighted code. |
| RENDER-07 | pass (#786) | The suspected bug is fixed: the `c++` fence gets a code bar with "c++" and "Copy". |
| RENDER-08 | pass | "Copied!" at 0.3 s, "Copy" at 2.3 s. Clipboard content not checked. |
| RENDER-09 | pass | A real click on the "Sheet1!B1" card-header link selects Sheet1!B1 (read back through Office.js). The tab stays idle and focus stays on the link. |
| RENDER-10 | pass (#786) | The suspected bug is fixed: a link to the deleted sheet "Gone" shows the error toast "Couldn't go to Gone!A1: The requested resource doesn't exist." with ×, and no unhandled rejection. |
| RENDER-11 | pass | A CSV-mode read of Sheet1!C5:E9 shows the grid with columns C–E, rows 5–9 and "Copy CSV". A default read shows the table without "Copy CSV" (as corrected). |
| RENDER-12 | pass | Offline: "Error: Failed to fetch". Stop mid-stream kept the text and ended with "Request aborted", with no banner. |
| RENDER-13 | pass | An injected `user-with-attachments` message shows two document tiles with truncated names and one image tile; clicks do nothing. |
| KEYS-01 | pass (#793, #786, #795) | The page changed on purpose. A web note says the browser keeps its tab chords and points to + and the tab menu. Rows: Enter, Enter (while streaming), Alt+Enter, Alt+↑, Shift+Tab (in the message box), Ctrl+O, ← → (with a tab focused), Shift+F10, /, ↑ ↓, F2, Esc. F6 and browser-reserved chords are gone (#793). |
| KEYS-02 | pass (#793) | web-keys: /shortcuts opened by command closes on the first Escape (it closes on the page it opened at, REVERT-16 rule), and focus goes to the composer. |
| KEYS-03 | pass (#793) | Toast "Wait for this tab to finish before running /shortcuts" (#793 wording). |
| KEYS-04 | pass (#789) | A busy-refused command chosen from the command menu with a real Enter keeps "/shortcuts" in the composer. |
| KEYS-05 | pass | web-keys: Ctrl+T opened a browser tab. Pi's tabs were unchanged and no keydown reached the taskpane. |
| KEYS-06 | pass | web-keys: Ctrl+W closed the Excel browser tab. Ctrl+Shift+T then gave a tab with no taskpane. |
| KEYS-07 | pass | web-keys: with a closed browser tab, Ctrl+Shift+T reopened it and nothing reached Pi. With none (CDP key), Pi reopened its tab ("Reopened: Chat 5"). |
| KEYS-08 | pass | web-keys: Ctrl+PageUp with a Pi tab focused switched the browser tab, with no key events in the taskpane and no change to Pi's tabs. |
| KEYS-09 | pass (#786) | web-keys: Ctrl+O shows "Details expanded (Ctrl+O)", then "Details collapsed (Ctrl+O)"; `pi.transcript.details-expanded` is 1 then 0. No Open dialog. |
| KEYS-10 | pass (#793) | It failed on `main` (the document was corrected for the web). Now Shift+Tab in the composer cycles the level and focus stays in the composer. |
| KEYS-11 | pass | web-keys: Escape in an idle composer focuses the active tab; F2 brings focus back, and the draft is kept. |
| KEYS-12 | pass | web-keys: ← and → on a focused tab switch tabs. |
| KEYS-13 | pass (#789) | Alt+↑ with nothing queued: toast "No queued messages to edit" (#789 wording). |
| KEYS-14 | pass (#793) | F6 is no longer listed (#793). Pi does not handle it; the page lost focus to the browser. |
| KEYS-15 | pass | Ctrl+Z in an empty composer with no Undo toast showing: no reopen, A1 still "keep", no toast. |
| KEYS-16 | pass (#789) | Responding tab, resume picker opened from the gear: Escape closes it and the turn keeps streaming. |
| KEYS-17 | blocked | Needs Excel desktop. |
| KEYS-18 | blocked | Needs Excel desktop (Mac). |
| KEYS-19 | blocked | Needs Excel desktop for Windows. |
| UTIL-01 | pass | Seven items in order: Settings, Extensions, Files, Rules, Resume session, Backups, Keyboard shortcuts. The menu is opaque (#793). |
| UTIL-02 | pass (#793) | Behavior changed (#793 menu model): focus moves to "Settings", ↓/↑ wrap, Home jumps, and Escape returns to the gear. On `a2b3f2d` and `e48d83c`, Tab closed the menu and dropped focus out of the taskpane (new bug, fixed in `8099cd6`). Now Tab closes to the gear. |
| UTIL-03 | pass | With the menu open, a real click on another tab closes the menu and activates that tab. |
| UTIL-04 | pass | Right-clicking a tab with the gear menu open shows only the tab context menu, with focus on "Rename tab…". Shift+F10 on a focused tab also opens it (#793), and Escape returns focus to the tab. |
| UTIL-05 | pass | Escape closes the menu while responding; Stop is still shown. |
| UTIL-06 | pass (#793) | Each item opens its view: Settings, Connections, Files, Rules, Resume session, Backups, Keyboard shortcuts. Each overlay takes focus (#793). |
| UTIL-07 | pass | Fresh browser context: "Resume session" toasts "No previous sessions" and opens no overlay. |
| UTIL-08 | blocked | Needs Excel for Mac. |
| LANG-01 | pass | "Language / 语言" with "English" and "中文（简体）· AI 翻译". |
| LANG-02 | pass (#796) | Picking Chinese: the toast is now in the current language, "Language changed. Reloading..." (the language only takes effect through the reload, #796), and the reload follows 0.86 s later. |
| LANG-03 | pass (#789, #795) | Changed on purpose: the draft ("草稿 langdraft") survives the language reload (#789), and tab numbers are kept, 对话 1, 3, 5, 6, 7, 8, 9 (#795). |
| LANG-04 | pass | Tabs read 对话 n, the bar reads 高 / 自动, and the disclosure banner and placeholder are Chinese. `<html lang="zh-CN">` (#796). |
| LANG-05 | pass (#786) | The Ctrl+O toasts are now translated ("已展开详情（Ctrl+O）"). `/yolo status` stays mixed: "自动 mode — Pi applies workbook changes immediately." |
| LANG-06 | pass (#796) | Back to English: toast "语言已更改，正在重新加载…", reload, `lang="en"`. |
| LANG-07 | pass (#796) | The suspected bug is fixed: switching while busy asks "Stop the response and switch language?". Escape cancelled it, the select went back to English, and the turn continued. |
| LANG-08 | pass (#796) | The suspected bug is fixed: with the `language` put made to throw, the English error toast "Failed to save language setting." appeared with ×. The select reverted, there was no reload, and the UI stayed English. |
| LANG-09 | pass | Fresh context, `/__pi-auth` blocked: on the welcome screen, "中文" reloaded into Chinese at once with no toast, and "English" switched back. |
| LANG-10 | pass | The toast "Dark mode: enabled" appears but the taskpane stays light: Office on the web reports a white `officeTheme`, and the OS is not dark here. Then "Dark mode: disabled". |
| LANG-11 | blocked | Needs Excel desktop. |
| PERSIST-01 | pass | After a reload: three tabs, the third active, with models GPT-6 Luna / GPT-6.1 Sol / GPT-6 Luna and levels Low / High / Min. A level change alone is only saved with the next message or model change (same as on `main`). |
| PERSIST-02 | pass (#795) | Changed on purpose: numbers are stable instead of renumbered. After closing Chat 2 and reloading: Chat 1, Chat 3, Chat 5. |
| PERSIST-03 | pass (#789, #795) | Changed on purpose: the draft "draft three" comes back after a reload (#789), and Ctrl+Shift+T after a reload reopens the closed tab ("Reopened: Chat 8", #795). |
| PERSIST-04 | pass (#795) | `Office.context.document.url` is "". The layout is stored under `workbook.tabLayout.v1.doc_instance:<token>`, with no `__global__` slot (#795, SESS-22). |
| PERSIST-05 | pass | "safe", then "yolo". |
| PERSIST-06 | pass | `navigator.storage.persisted()` is false (#796 asks, Chromium declines in the frame). |
| PERSIST-07 | pass (#794) | The suspected bug is fixed: 17, then 18, backups survived several reloads. |
| PERSIST-08 | pass (#795) | With the model request held in the page, the reload brings back the tab with its unanswered "PERSIST08c…" message, idle. (A request held with CDP `Fetch` blocks the reload itself until it is released.) |
| PERSIST-09 | pass (#789) | The suspected bug is fixed: with session puts made to throw, the toast "Couldn't save this chat: injected" appears with ×. |
| PERSIST-10 | pass (#789) | Changed on purpose: after auto-compaction (11%) and after a popover /compact (18%), a reload keeps the compacted context. |
| PERSIST-11 | blocked | Needs Excel desktop. |
| PERSIST-12 | blocked | Needs a signed-in OneDrive workbook in Excel desktop. |
| MULTI-01 | pass | Pane B opened on the same link with pane A's tabs. |
| MULTI-02 | pass (#794) | Changed on purpose (#794: one mode for every pane): B set to Auto, A showed "Auto" at once, and the stored value is "yolo". |
| MULTI-03 | pass | Last write wins: B's close, and then B's new tab, replaced the stored layout (1, then 2 ids). Side note: B's new tab was numbered "Chat 3" while A had a different "Chat 3". |
| MULTI-04 | pass | A tab closed in B, then Ctrl+Shift+T in A: "No recently closed tab". |
| MULTI-05 | pass | With B's `Excel.run` held 10 s (58.6–71.8 s), A's write calls ran at 60.7–62.7 s. A never showed "Waiting". D1 = 51, D2 = 52. |
| MULTI-06 | pass (#794) | The suspected bug is fixed: after reloading B, A's 32 backups, including D1 and D2, are still stored. |
| MULTI-07 | pass (#794) | The suspected bug is fixed: the audit log holds both panes' D1 and D2 writes, and both backups are kept. |
| MULTI-08 | pass | After the writes, both panes read D1:D2 as 51, 52. No lag was seen at this delay. |
| MULTI-09 | pass (#794) | The suspected bug is fixed: ALPHA in A and BRAVO in B on one session. BRAVO was saved as a new conversation, the original kept ALPHA, and later turns (CHARLIE, DELTA) stayed in their own sessions. The "saved as a new conversation" toast was not caught by a 9 s poll. |
| MULTI-10 | blocked | The orb's login comes from `/__pi-auth` (222 h left), so the browser-grant refresh path is reached only by replacing the shared grant. That was not done, to protect it. #795 covers single-flight refresh with unit tests. |
| MULTI-11 | blocked | Needs Excel desktop. |
| PROVIDER-01 | pass (#789) | The suspected bug is fixed: offline shows "Error: Failed to fetch" and the banner "You're offline. Reconnect, then send your message again.", not the CORS banner. |
| PROVIDER-02 | pass (#789) | Changed on purpose: with the proxy enabled and stopped, the banner reads "Couldn't reach the provider through the proxy at https://localhost:3003. Check that the proxy is running, then retry." |
| PROVIDER-03 | pass (#793) | The warning "The proxy is turned on but isn’t running, so some providers and web tools won’t work." appeared 8 s and 29 s after the stop. It went ~31 s after the restart, and at once when the proxy was turned off in Settings (`8099cd6`/`e48d83c` re-check). |
| PROVIDER-04 | pass (#793) | Wrong OpenAI (API) key on GPT-4.1 mini: "Error: OpenAI API error (401): …" and the banner "Couldn’t get a response: OpenAI API error (401): …". No dialog. |
| PROVIDER-05 | pass (#789, #793) | Changed on purpose (per-tab errors): the request was delayed 5 s and the user switched away. The active tab got no banner. The background tab's label became "Stopped with an error" (#793), and its banner shows when switching to it. |
| PROVIDER-06 | pass (#789) | Changed on purpose: the banner follows its tab. It is hidden on other tabs and back on its own. |
| PROVIDER-07 | pass (#789) | With the key removed, "Connect OpenAI (API)" opens with focus in the key field. Closing it gives "Error: Provider is not configured: openai" and the banner "API key required for openai." |
| PROVIDER-08 | pass (#789) | Changed on purpose: two substituted 503s gave "Retrying in 2s (attempt 1 of 3)… Provider hiccup", then 4 s (attempt 2), then "OK". |
| PROVIDER-09 | pass | A substituted overflow on an 18% session gave "Compacting to free up context", "Summarized 2 messages", and a retried request answering "OK". On a small session it gave "Nothing to compact" and "Context window exceeded. Run /compact…", with no retry promised (#789). |
| PROVIDER-10 | pass (#789) | The suspected bug is fixed: a 400 "Request was cancelled upstream" shows the banner "Couldn’t get a response: Request was cancelled upstream". |
| PROVIDER-11 | pass on `8099cd6` | It failed on `e48d83c` (regression, #789): after the proxy was stopped mid-stream, #789's retry dropped the partial answer. On `8099cd6` (#789 a66868d), the partial text (325 chars) is kept and ends with "Error: Failed to fetch", plus the proxy banner. |
| PROVIDER-12 | blocked | Needs Excel for Mac. |
| HOST-01 | pass | OfficeOnline; ExcelApi 1.16–1.20 true, 1.21 false. |
| HOST-02 | pass (#795) | The suspected bug is fixed: `/copy` gives "Copied to clipboard" (execCommand fallback). |
| HOST-03 | pass | The proxy banner's "Show how to start it", then "Copy command", selects "npx pi-for-excel-proxy" (the clipboard is blocked). |
| HOST-04 | pass | See KEYS-05 to KEYS-08. |
| HOST-05 | pass (#793) | Changed on purpose: on the web, Shift+Tab in the composer cycles the level and no longer moves focus. From other controls it moves focus back without cycling. |
| HOST-06 | pass (#794) | `isDirty` is false, and backups are no longer cleared on opening (see PERSIST-07). |
| HOST-07 | pass (#794) | The suspected bug is fixed: with the document token removed, Pi wrote B17. The backup count went from 17 to 18 and stayed at 18 for more than 15 s. |
| HOST-08 | blocked | Needs Excel desktop (Windows half; Mac recorded on `main`). |
| HOST-09 | blocked | Needs a OneDrive workbook with AutoSave in Excel desktop. |
| HOST-10 | blocked | Needs Excel desktop (#796's Mac download fix). |
| HOST-11 | pass, with a difference | In the Excel frame, `document.featurePolicy.allowsFeature("loopback-network")` is false and the permission is "denied"; there is no prompt. The dev build is served from localhost, so #796's "browser blocks the proxy" banner is deliberately not used (`isLoopbackHostname`). The proxy works here only because the orb's Chrome turns Local Network Access checks off. |

## Redesign at ~340 px, keyboard and live regions

These checks ran in the 349 px Excel on the web pane (screenshots inspected, kept in the orb under `.amp/in/artifacts/ui-verify/`).

- **Layout**: the notice area sits under the tabs, most severe first: run error (red), proxy warning (amber), disclosure (info). Each banner has ×. The dock holds the working line and the input in one card. The status bar stays on one row; the model name truncates to "GPT-6 Luna · Chat…" and the limit hides. The mode popover, the opaque gear menu, the resume picker and Settings all fit. The Settings card scrolls under a pinned header.
- **Toasts**: they stack above overlays. A toast over Settings hits `#pi-toast` with `elementFromPoint`, at z-index 1000. Error toasts have ×.
- **Tabs (8099cd6)**: with 7 tabs, titles read "Chat N" (52–55 px tabs, full-width titles). There is no native scrollbar. The active tab is scrolled into view, and ‹ › scroll to the ends. On `e48d83c`, every tab was 60 px with a 17 px title ("C…"), and a tab with a marker showed only "(" and a dot.
- **Keyboard-only, real keys**:
  - Composer → Send → model → thinking → context → mode with Tab, and back with Shift+Tab (on `e48d83c`/`8099cd6`; on `a2b3f2d`, Shift+Tab on Send cycled thinking).
  - Menus and popovers open with focus on the first or checked item, ↑/↓ wrap, Enter chooses, Escape returns to the trigger, and Tab closes to the trigger (`8099cd6`).
  - Overlays (/resume, approvals) take focus. Tab and Shift+Tab cycle inside, and Escape closes and returns focus to the composer (`8099cd6`).
  - The approval dialog: focus is on the card, with Stop Pi / Decline / Allow once. Everything behind it is inert. Stop Pi withdraws the request ("Request aborted", B30 unchanged), and Escape declines.
- **Live regions**:
  - The toast region is `aria-live=polite`, and error toasts are `role=alert` (recorded on show).
  - `.pi-announcer` is polite and atomic ("Pi finished in “Chat 3”.").
  - The disclosure banner is `role=status`.
  - On `8099cd6`, both live regions stay live (not inert) while an approval is open; on `e48d83c` they were made inert.
  - Workbook approvals are still not announced (minor, see below).

## Regressions and new bugs

Fixed during the pass (each reported to its owner when found):

1. **A tab waiting for the workbook lock never shows it (STATUS-14, STATUS-15).** Regression, #794, on `a2b3f2d`; fixed in `e48d83c`.
   - Repro: three tabs. Delay `Excel.run` 30 s, write in Chat 3, and once the bar shows "Changing workbook", remove the delay. Write in Chat 2: it waits about 26 s with no "Waiting for workbook" and no tab badge.
   - Cause: `src/workbook/coordinator.ts:229-238` emitted `queued` before pushing the operation, so the subscriber in `src/taskpane/init.ts:1904-1916` read a snapshot without it.
2. **Shift+Tab on the Send or attach button cycled thinking and kept focus there.** New bug, #793; fixed in `e48d83c`.
   - Repro: from the thinking button, press Shift+Tab twice to reach Send, then Shift+Tab again: Low → Medium.
   - Cause: `src/taskpane/keyboard-shortcuts.ts:88-96` (`isInEditor` counts all of `.pi-input-card`), used at `:307`.
3. **← or → inside a status popover switched Pi tabs.** New bug, #793; fixed in `e48d83c`.
   - Repro: two tabs, open the thinking popover, press →.
   - Cause: `src/taskpane/keyboard-shortcuts.ts:285-297`; `src/ui/menu-keyboard.ts:65-86` did not consume horizontal arrows.
4. **Tab in the gear menu or a status popover dropped focus out of the taskpane.** New bug, #793, on Excel on the web only; fixed in `8099cd6`.
   - Repro: click the gear, press Tab. activeElement is BODY and `document.hasFocus()` is false. A standalone page was fine.
   - Cause: `src/ui/menu-keyboard.ts:53-57`, together with Office on the web's window Tab handling.
5. **The overlay focus trap let the first Tab from the card leave the add-in frame.** New bug, #793; fixed in `8099cd6`.
   - Repro: /resume or an approval, press Tab. Focus leaves the frame, and returns on the second Tab.
   - Cause: `src/ui/overlay-dialog.ts` trap.
6. **Tabs shrank to 60 px with unreadable titles.** New bug, #793 82078ec, on `e48d83c`; fixed in `8099cd6`.
   - Repro: 7 tabs at 349 px. Every tab reads "C…", and a tab with a marker shows only "(" and a dot.
   - Cause: `src/ui/theme/components/tabs.css:96-105` (`flex-shrink: 1; min-width: 60px`).
7. **An open approval made the screen-reader live regions inert.** New bug, only in the combination of #794 and #793; fixed in `8099cd6` (#794 515072d).
   - Repro: Confirm mode, open an approval. `.pi-announcer` is inert.
   - Cause: `src/ui/approval-dialog.ts:201-204` made every body child inert.
8. **A mid-stream connection drop lost the partial answer (PROVIDER-11).** Regression, #789, on `a2b3f2d`/`e48d83c`; fixed in `8099cd6` (#789 a66868d).
   - Repro: GPT-6 Luna, thinking Off, ask for 400 lakes, and stop the proxy after ~400 chars. The stream fails with ERR_INCOMPLETE_CHUNKED_ENCODING, a retry goes out 2 s later and fails with ERR_CONNECTION_REFUSED, and the final message has `content: []` and shows only "Error: Failed to fetch".
   - Cause: `retryTransientFailure` in `src/taskpane/action-queue.ts:240-270` dropped the failed response before retrying.

Still open on `8099cd6` (all minor):

9. **The hover × draws over the tab title.** #793.
   - Repro: 7 tabs at 349 px, scroll to the start, hover "Chat 1" (52 px). It renders "Cha×1".
   - Cause: `src/ui/theme/components/tabs.css:255-273`; the × on an inactive tab overlays the title instead of taking space.
10. **Workbook approvals are not announced.** #793.
    - Repro: Confirm mode, have Pi write a cell. The approval opens, and `.pi-announcer` keeps its previous text ("Pi finished.").
    - Cause: only the runtime-tool path calls `announce(t("announce.approval"))` (`src/taskpane/init.ts:1028`); the workbook path at `:1122` does not. Low impact, because the dialog takes focus.
11. **The busy model picker stays read-only after the turn ends.** #795.
    - Repro: open the model control during a long answer and leave it open until the answer finishes. All rows stay `aria-disabled` and the note stays; the user must close and reopen the picker.
    - Code: the busy state is drawn once when the picker opens (`src/ui/model-selector-dialog.ts:261`).

Observed, not counted as bugs: across two panes, stable tab numbers can collide (B's new tab became "Chat 3" while A had a different "Chat 3", #795). Workbook approvals were often preceded by the model asking "Please confirm…" in text; that is agent behavior, not UI.

## Desktop items to hand to the Mac

STATUS-21, KEYS-17, KEYS-18, KEYS-19 (Windows), UTIL-08, LANG-11, PERSIST-11, PERSIST-12 (needs OneDrive sign-in), MULTI-11, PROVIDER-12, HOST-08, HOST-09 (OneDrive with AutoSave), HOST-10.

Worth rechecking there with this build:

- #793's overlay and menu focus in WKWebView (R2 in the design).
- Shift+Tab staying in the composer (KEYS-17).
- #796's blob-download path for `/export` and `/backup` (HOST-10).
- The tab-rename dialog in place of `window.prompt` (HOST-08).
