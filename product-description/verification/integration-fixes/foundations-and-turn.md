# Integration verification: foundations and turn

This file reruns every item in [../foundations-and-turn.md](../foundations-and-turn.md) against the combined fix build, branch `integration-fixes`. That branch is `main` plus PRs #786–#799 (see `INTEGRATION.md` on the branch). The full pass ran at `a2b3f2d` on 27 October 2026. After the branch was force-pushed to `e48d83c`, the dev server was restarted from that commit and the items touched by the new fixes were rerun: toasts against the Not sent rows, the "+" button and tab strip, the proxy banner, compaction toasts and saving, Stop and close, the selection note and COMPOSE-11. The notes say where a result is from `e48d83c`. The branch was then force-pushed to `8099cd6`, and the items it changed were rerun there: COMPOSE-11, SEND-03, SEND-13, the compaction toasts, the tab strip, /files Escape, the tab and slash menus, and Stop during a tool call under #792's new Stop rules (see "Rechecked on `8099cd6`"). On `cea7588` the tab-close fix and the slash menu under toasts were rechecked (see "Rechecked on `cea7588`").

Expected behavior is the checklist as amended by the fix PRs and `goal.md` "Fixes in flight". Where a PR deliberately changed behavior, the PR wins, and the result names it ("pass (#NNN changed)"). "pass (wording)" means the behavior matches and only the text changed.

How it ran:

- **web**: Pi in Excel on the web through `node scripts/orb/excel-web-session.mjs --check` (host Excel, platform OfficeOnline, check ok), with the dev server serving this checkout. The taskpane (349×622) was driven over CDP on its own target with Playwright: real key presses and mouse events, DOM and agent state read back, and a MutationObserver recording the Stop button, working indicator, queue rows, banners, toasts and tabs.
- **web-grid**: the Excel grid was clicked, then keys were typed into it.
- **web-inject / web-observe**: the active tab's agent `streamFunction` was wrapped from the console, either to throw once after a delay (provider failure, context overflow) or to record each request's messages (auto-context notes and reasoning level). A 429 and a 503 were returned for one request with CDP route interception.
- **web-slow**: `Excel.run` in the taskpane was wrapped with an 8 s delay.
- **Compaction** used ~30–50k-token tabs built from paged reads of a 3,000-row Data sheet. The context window was shrunk in memory (`agent.state.model.contextWindow`) for the meter colours, pre-send compaction, compaction between responses and STOP-13.
- **Reload** used `location.reload()` in the taskpane frame; CDP `Page.reload` refuses non-top-level targets. Once, Office.js lost its session after several reloads; rerunning the session script reopened Pi.
- **Model turns** ran on GPT-6 Luna (GPT-6.1 Luna is not in the registry), with two control turns on GPT-6.1 Sol. The shared `PI_EXCEL_WEB_AUTH_JSON` login was never logged out or refreshed.

**Result at `8099cd6`: 147 items. 143 pass, 0 fail, 4 blocked (desktop only, skipped).** At `a2b3f2d`/`e48d83c` it was 141 pass and 2 fail (COMPOSE-11, SEND-03).

- Every item that failed on `main` now passes or behaves as its owning PR describes, except COMPOSE-11. #789 claimed that fix, and it still fails on this build (fixed later in #789 @ 4354892).
- One item that passed on `main` regressed: SEND-03, because a steer now turns following back on (#789; fixed in 4354892).
- The pass found nine new bugs, five of them fixed by `e48d83c`. Still open on `e48d83c`: the tab scroll arrows, /files Escape focus, and the "Compacting…" toast after a stopped compaction (all #793), plus COMPOSE-11 and SEND-03 (#789, fixed after this build). A possible regression in SEND-13 (#789) is waiting for its owner's answer.
- On `8099cd6` all of these pass: COMPOSE-11, SEND-03 and SEND-13 (#789 @ a66868d) and the three #793 items. One new bug: on a narrow inactive tab, the × that appears on hover covers the right half of the title, so clicking the title can close the tab (#793, item 10). Fixed on `cea7588`.

## Results

| ID | Result | Note |
| --- | --- | --- |
| TASKPANE-01 | pass | After reload: focus in the composer; status bar "GPT-6 Luna", "High", "1%/272k", "Auto". |
| TASKPANE-02 | pass (#793 changed) | The run-error banner is now in the notice area under the tab strip (banner top 53 px, tabs at 0), closable, not above the header. |
| TASKPANE-03 | pass (#789: per-tab drafts) | "SHARED-DRAFT" in Chat 1; "+" opened Chat 2 with an empty composer; back on Chat 1 the draft was there; still there after a taskpane reload. |
| TASKPANE-04 | pass | Chat 2 responding → click Chat 1: no indicator, send button, rotating placeholder; Chat 2 kept its busy dot ~13 s until done. |
| TASKPANE-05 | pass (#793 changed) | /files while responding now takes focus (filter field). On e48d83c the first Escape dropped focus to <body> and left the overlay open (minor new bug, #793; fixed on 8099cd6, where focus moves to the dialog card); the second closed it; the turn kept streaming. A third Escape stopped the turn. |
| TASKPANE-06 | pass | 8099cd6: the first Escape moves focus from the filter field to the dialog card (overlay open), Tab stays inside the overlay over 12 presses, and the second Escape closes it with focus back in the composer. On e48d83c focus went to <body> after the first Escape; second Escape closed it and focus returned to the composer. |
| TASKPANE-07 | pass (#793 changed) | Shift+Tab in the composer stepped the level each press and focus stayed in the composer; from the thinking button Shift+Tab did not change the level and moved focus back to the model button. |
| TASKPANE-08 | pass | Focus on a tab, F2: focus in the composer. |
| TASKPANE-09 | pass (#793 changed) | The Ctrl+O toast "Details expanded (Ctrl+O)" sits just above the composer dock (y 485–518 in a 622 px pane), not under the header: #793 moved toasts above the dock. |
| TASKPANE-10 | pass (#793 changed) | The disclosure bar no longer has "Customize" with four switches; it has "Got it", "Review in Settings" and ×. "Review in Settings" opened the Settings shell with focus inside it. |
| TASKPANE-11 | pass (#793 changed) | The dead switches are gone (replaced by "Review in Settings"), so the suspected bug no longer applies. |
| TASKPANE-12 | pass (#793) | A toast shown while Settings was open drew above the card: elementFromPoint at the toast's centre hit the toast. Escape closed Settings. |
| TASKPANE-13 | pass (#793 changed) | With the proxy off in Settings, stopping the proxy service showed no banner (by design). With the proxy on, stopping the service showed "The proxy is turned on but isn't running, so some providers and web tools won't work. Show how to start it ×" about 29 s later. New bug: turning the proxy on while it is already down shows no banner. |
| TASKPANE-14 | pass | desktop (macOS), `8099cd6`: after a real click on "Open Pi", keys typed at once ("hi") went into the composer. |
| TASKPANE-15 | pass | Taskpane frame set to 490 px: composer limit 137 px = 28%. The π mark was not on screen in this answered tab at either height, so only the composer limit was checked. |
| TURN-01 | pass | Enter sent; message in transcript, composer cleared, Stop, busy dot, "Working…" within ~10 ms. A "Queued" row flashed for ~8 ms before the turn started (pre-send compaction check, #789). |
| TURN-02 | pass | Indicator, Stop and dot went at turn end; placeholder rotating again. |
| TURN-03 | pass | Rows "Steer Then say BANANA." and "Follow up Then say MANGO." (labels now sentence case with Edit and ×, #789). |
| TURN-04 | pass | BANANA answered first after the response, then MANGO; strip empty at the end. |
| TURN-05 | pass (#789) | Stop button with steer+follow-up waiting: two "Not sent" rows (Edit, ×), toast "Stopped · 2 queued messages not sent, kept above the composer"; next turn "Reply with just OK." answered "OK", nothing sent unasked. |
| TURN-06 | pass (#789) | Escape with steer+follow-up waiting and draft "DRAFT-TEXT": same Stop as the button: two Not sent rows above the composer, draft kept, toast "Stopped · 2 queued messages not sent…", "Request aborted" in the transcript. |
| TURN-07 | pass (#792 changed) | Stop during a 12 s execute_office_js no longer waits for the tool: #792's settleOnStop settled the call and the turn ended 0.79 s after Stop. 8099cd6: a read_range with Excel.run delayed 12 s settled 2.6 s after Stop ("Stopped by the user before Excel answered.", red card, "Request aborted"). |
| TURN-08 | pass (#789) | Ended with "Request aborted", not "signal is aborted without reason"; no banner. |
| TURN-09 | pass (wording) | /shortcuts while responding: refused with the toast "Wait for this tab to finish before running /shortcuts" (wording changed from "Can't run /{command} while Pi is busy"); nothing opened. /model is now busy-allowed (#795), so it was not usable for this item. |
| TURN-10 | pass | /compact while responding: row "/compact ×"; compaction ran right after the turn. |
| TURN-11 | pass | Chat 2's answer complete on return (300 lines, stop). |
| TURN-12 | pass | Background Chat 4 responding, its ×: dialog "Stop and close this tab? … Pi is still responding in this tab. Stop and close it? Keep open / Stop and close". "Keep open" left it running. |
| TURN-13 | pass (#795 changed) | While busy, /model and the status-bar model both open the picker read-only: note "Pi is working in this tab. Change the model when it finishes, or press Stop.", 7 rows aria-disabled; clicking one left the picker open and the model unchanged; Escape closed the picker and the turn kept streaming. |
| TURN-14 | pass (wording) | Injected failure: "Error: Simulated provider failure" in the transcript; banner "Couldn't get a response: Simulated provider failure ×" (wording changed from "LLM error: …"). |
| TURN-15 | pass (#789 changed) | Failure injected with a 4 s delay, steer and follow-up sent within it: after the failure both became Not sent rows; the agent's queues were empty, and the next turn answered only "OK". |
| TURN-16 | pass (#789 changed) | Banners are per tab: Chat 4 showed none while Chat 5 had one; Chat 4 failed in the background while Chat 5 was active (Chat 5 kept its own banner, Chat 4 got a failed marker); back on Chat 4 its own banner "…Simulated provider failure B" showed. |
| TURN-17 | pass (#795 changed) | Reload with answered Chat 5 and new Chat 11 both responding: Chat 5 ends with the user message and no answer; Chat 11 is no longer empty, it keeps its user message (#795 saves from the first user message). |
| TURN-18 | pass (#789 changed) | /compact while idle shows the Stop button; Escape stopped it: toast "Compaction stopped", nothing compacted. Next to it the stale "Compacting to free up context" toast was still up (new bug, #793). |
| TURN-19 | pass | Prompt sent during compaction: row "Queued Reply with just QUEUED-OK."; sent after "Summarized 1 message" and answered QUEUED-OK. |
| TURN-20 | pass (wording) | Confirm mode, Pi asked to write F1: dialog "Allow this change to the workbook? … Write Sheet1!F1 … Confirm mode: Pi asks before each change… Stop Pi / Decline / Allow once" with the rest of the taskpane inert; Escape declined; result "The user declined this workbook change in the approval dialog (Confirm mode)."; Pi answered "DECLINED"; F1 empty. |
| TURN-21 | pass | Shift+Tab while responding: toast "Thinking: Extra high (next turn)". |
| TURN-23 | pass (wording) | Excel.run delayed 8 s, Chat 4 and Chat 5 each asked to set a cell with execute_office_js 0.2 s apart: Chat 5 showed the badge "Waiting" (tooltip "Waiting for another tab to finish changing the workbook") and the status bar "Waiting for workbook" while Chat 4 wrote, then "Changing workbook" for its own write; J1 = 61 and J2 = 62. The badge now says "Waiting" instead of "Locking…". |
| TURN-24 | pass | With Excel.run delayed 8 s, while Chat 5's write held the lock its × was disabled with the tooltip "Wait for workbook changes to complete"; enabled again after the write. Status bar showed "Changing workbook". |
| TURN-22 | pass | desktop (macOS), `8099cd6`: Auto mode, Pi's `fill_formula` of `=ROW()*2` over `Sheet1!F1:F1000000`; Stop pressed about 1 s after the card appeared. The turn ended "aborted" and the fill handed to Excel still completed (F1 = 2, F1000000 = 2000000); the card reads "Filled Sheet1!F1:F1000000 — 1000000 changed, no backup". Same as on `main`. |
| COMPOSE-01 | pass | Shift+Enter: "line one\nline two" in the composer, nothing sent. |
| COMPOSE-02 | pass | Alt+Enter while idle sent both lines as one message. |
| COMPOSE-03 | pass | Idle placeholders changed every ~8 s in the order "Have Pi edit…" → "Summarize…" → "Ask about…" → "Type / for commands…"; fixed "Redirect reply (↵) · New question (⌥↵)" while responding (TURN-01). |
| COMPOSE-04 | pass | Escape while idle: focus on the active tab ("Open tab Chat 5"), draft "KEEP" kept; F2 back to the composer. |
| COMPOSE-05 | pass (#789) | Typed /shortcuts (menu showed it), Enter while responding: toast, draft "/shortcuts" kept. |
| COMPOSE-06 | pass (#789 changed) | /zzz + Enter while idle: nothing sent, draft "/zzz" kept with no added newline, toast "No command /zzz. Start with a space to send it as a message". |
| COMPOSE-07 | pass (#789 changed) | Send button with "/compact" runs the command (toast "Compacting to free up context"); no user message carrying "/compact" was added. |
| COMPOSE-08 | pass (#789) | As TASKPANE-03. |
| COMPOSE-09 | pass | As TASKPANE-04. |
| COMPOSE-10 | pass (#789) | As TURN-06. |
| COMPOSE-11 | pass (8099cd6) | 8099cd6 (#789 @ a66868d): five typed Shift+Enter lines and clearing the draft kept following on, gap 0. Failed on a2b3f2d and e48d83c: with the transcript at the bottom, three typed Shift+Enter lines turned following off (gap 59–98 px). #789 claims this fixed. Cause: pi-input's `height: auto` autosize (src/ui/pi-input.ts:182) briefly collapses the composer, Chrome clamps scrollTop, and the scroll listener reads that as the user scrolling up. #789's owner reproduced it and fixed it in #789 @ 4354892, after this build. |
| COMPOSE-12 | pass | CDP Input.dispatchDragEvent with a real file on the composer: toast "Imported 1 file into Files.", draft "KEEP-DRAFT" kept (drag-over hint not captured, as before). |
| COMPOSE-13 | pass | desktop (macOS), `8099cd6`: Simplified Chinese Pinyin, keys posted at the HID event tap. With the candidate bar "1 你 2 尼 3 泥 4 腻 …" open over the composer (screenshot), Return committed "ni" and left it in the composer; nothing was sent (23 messages before and after). Two earlier tries in which the input method did not engage (no candidate bar) sent "ni" and are not counted. The second-Return send was not rechecked here (it passed on #789 alone). |
| COMPOSE-14 | pass | A paste event carrying a PNG added nothing to the draft and imported nothing. |
| COMPOSE-15 | pass | 40 lines inserted in a 622 px pane: composer stopped at 248 px (40%) and scrolls. |
| SEND-01 | pass | "Working…" for 1.25 s, then a phrase ("Finagling…"); phrases and hints rotated ("Adjusting for seasonality…", "Discombobulating…"; hints "⇧Tab to adjust reasoning depth", "Type / for commands", "Press Enter to redirect Pi"). Hint timing was not measured. |
| SEND-02 | pass | Gap 30, 0, 34, 25, 21, 19 px → following on, on, off, off, off, on. |
| SEND-03 | pass (8099cd6) | 8099cd6: a send turns following on; scrolled up 400 px while streaming, a steer left following off (gap grew to 445 px). Failed on a2b3f2d and e48d83c: sending resumes following (401 px up → on after send), but a steer now resumes it too (400 px up, steer → following on, gap 0); on main a steer left it off. Not deliberate: #789's `_onSend` set `_autoScroll = true` for every submit. Fixed in #789 @ 4354892, after this build. |
| SEND-04 | pass (#789) | With steer and follow-up waiting and following on, two Alt+↑ presses restored both into the composer; Enter sent one steer with both lines; following stayed on throughout. |
| SEND-05 | pass (#789 changed) | While compacting (not responding): pill "Compacting context…" with "Send messages and Pi will see them after compaction"; the Stop button shows instead of send, since compaction can now be stopped (#789). |
| SEND-06 | pass | Stop 150 ms after send: "Request aborted" alone. |
| SEND-07 | pass (wording) | As TURN-14. |
| SEND-08 | pass | Sending in Chat 4 cleared its banner at once. |
| SEND-09 | pass (#789 changed) | As TURN-16: a background failure shows its banner when you switch to that tab. |
| SEND-10 | pass (#795 changed) | Tab model set in memory to Anthropic with no key: the message is in the transcript, "Connect Anthropic" opens over it with focus in its key field, tab responding; Escape aimed at the composer did not close it (focus is trapped in the dialog). Its × ended the turn with the banner "API key required for anthropic." (no longer \"Provider is not configured\"). |
| SEND-11 | pass (#789 changed) | Window shrunk to 18k (11.7k used): Enter started pre-send compaction and the message showed at once as "Queued Reply with just HELD-OK." with Edit/×; it was sent after the summary and answered "HELD-OK". |
| SEND-12 | pass (#789) | Switched to idle Chat 1 while Chat 2 compacted: the label left with the switch (one 29 ms frame during the re-render); Chat 1 showed no label and the send button, Chat 2 kept its busy dot. The toasts are still global. |
| SEND-13 | pass (#789 changed) | Proxy service stopped before sending: "Error: Failed to fetch" with banner "Couldn't reach the provider through the proxy at https://localhost:3003. Check that the proxy is running, then retry." (#789's own message, not the CORS one). Stopped mid-stream: #789 treats the network error as transient, drops the failed response and retries; with the proxy still down the turn ended "Error: Failed to fetch" with the same banner, so the partial answer was no longer kept on e48d83c (on main it was). 8099cd6 (#789 @ a66868d): the proxy stopped mid-stream, and the partial answer ("…forty-eight fort") stayed with stopReason error, "Error: Failed to fetch" and the proxy banner. Model traffic used the proxy only after Settings → Proxy was switched on; it was off in this profile. |
| SEND-14 | pass (#789 changed) | A 429 injected over CDP for one request on the ChatGPT login maps to "You have hit your ChatGPT usage limit." and is not retried (pi-ai does not mark it retryable). A 503 injected the same way was retried: "Retrying in 2s (attempt 1 of 3)… Provider hiccup", then "OK". |
| TOOL-01 | pass (#786 changed) | Three parallel reads: cards finish one at a time (inprogress×3 → complete,inprogress,inprogress → complete×2,inprogress → complete×3), group goes "2 reads" then "3 reads". |
| TOOL-02 | pass (#786) | Response drawn once: while the tools ran the cards were only in the list, none in the streaming container. |
| TOOL-03 | pass | Headers "Read Sheet1!A1:B3", "Read Sheet1!C1:D3", "Read Sheet1!A4:D6", collapsed (aria-expanded=false). |
| TOOL-04 | pass | Group "3 reads". |
| TOOL-05 | pass (#786) | "3 reads" collapsed: 0 of 3 cards visible; click → 3 visible, aria-expanded=true; click → hidden again. |
| TOOL-06 | pass (#786) | Live group "2 reads" (expanded) became "3 reads" (collapsed): 3+ start collapsed, 2 expanded, only user-chosen states remembered. |
| TOOL-07 | pass (#786 changed) | Ctrl+O first press: toast "Details expanded (Ctrl+O)", all 8 cards and both groups expanded; second press "Details collapsed (Ctrl+O)", all collapsed. |
| TOOL-08 | pass (#792 changed) | The card turned red (data-state=error) with "Stopped by the user while Excel was still working on this change. The outcome is unknown…", then "Request aborted". Not "waited": #792 settles stopped calls after 500 ms. |
| TOOL-09 | pass | The declined write card turned red (data-state=error) and the turn continued. |
| TOOL-10 | pass | The steer stayed in the strip ("Steer Then say BANANA.") while the tool ran. |
| TOOL-11 | pass (#786) | 12 switches between Chat 4 and Chat 5: Chat 5 always had exactly 15 cards and groups "3 reads,3 reads". |
| TOOL-12 | pass | Write card "Edit Sheet1!G1 — 1 changed"; body "Changes (1)" with Sheet1!G1 "∅" → "1", "Explain these changes", backup id in the result. |
| TOOL-13 | pass (#786) | fetch_page on ftp://example.com/file.txt: result isError=true ("Error: Only http(s) URLs are supported."), card data-state=error (red), no banner. |
| TOOL-14 | pass | Clicking the "Sheet1!C1:D3" link in a Read card selected Sheet1!C1:D3 in Excel; no card toggled. (A pointer click at the link's centre landed on the next card's header in the CDP geometry; the link was fired with element.click().) |
| TOOL-15 | pass | With #787's paging a 3,000-row read returns a 500-row page under the limit, so a 500-row long-text read (Data!D1:E500) was used: "[Output truncated: showing first 124 of 504 lines (50.0KB limit); limits: 2,000 lines / 50.0KB; full output saved to Files workspace: .tool-output/…read_range….txt]". |
| QUEUE-01 | pass | As TURN-03. |
| QUEUE-02 | pass | As TURN-04. |
| QUEUE-03 | pass (#789 changed) | Alt+↑ with steers waiting edits only the newest row ("x" back in the composer, the other rows stay); no toast. |
| QUEUE-04 | pass (#789) | Enter on the restored text gave one row "Steer Then say BANANA. Then say MANGO.". |
| QUEUE-05 | pass (wording) | Idle Alt+↑ with nothing waiting: toast "No queued messages to edit". |
| QUEUE-06 | pass | Only /compact queued: Alt+↑ toast "No queued messages to edit"; the "/compact ×" row stays. |
| QUEUE-07 | pass (#789) | As TURN-05. |
| QUEUE-08 | pass (#789) | After a reload both Not sent rows were still there; Alt+↑ put only the newest ("Then say MANGO.") back above the draft "DRAFT-TEXT"; × removed the other. |
| QUEUE-09 | pass (#789) | As TURN-15. |
| QUEUE-10 | pass (#789) | As COMPOSE-05. |
| QUEUE-11 | pass | As TURN-19. |
| QUEUE-12 | pass (#789) | Steer sent during the tool, then Stop: not added to the transcript; kept as a Not sent row with the toast "Stopped · 1 queued message not sent…". |
| QUEUE-13 | pass (#789 changed) | A 60-character steer is no longer cut to 47 characters: the row holds the full text (ellipsis by CSS only) with the full text as its title on hover. Label "Steer" in sentence case. |
| QUEUE-14 | pass | Two identical steers "x" showed two rows. |
| QUEUE-15 | pass (#789 changed) | Steer waiting behind a tool call, switched away, closed the tab with "Stop and close", /reopen: the steer was never sent and came back as a "Not sent" row (saved with the session); no draft. |
| QUEUE-16 | pass (#789) | Not sent rows survive a reload by design and were never delivered unasked (TURN-05's next turn answered only "OK"). |
| STOP-01 | pass | Stop after ~400 chars of text: text stops at "forty-one", "Request aborted", send button back, no banner. |
| STOP-02 | pass | Escape while streaming stopped with "Request aborted" (TURN-06). |
| STOP-03 | pass | As TASKPANE-05. |
| STOP-04 | pass | Thinking popover opened by a real click took focus; Escape closed it, focus back on the thinking control, turn still streaming. |
| STOP-05 | pass (#793 changed) | While responding, Shift+F10 opened the tab menu; Escape closed it (focus back on the tab) and the turn kept streaming. The menu now claims Escape. |
| STOP-06 | pass (#789) | As TURN-08. |
| STOP-07 | pass | As TURN-20. |
| STOP-08 | pass (#789) | Stop button and Escape now treat waiting items the same (TURN-05, TURN-06). |
| STOP-09 | pass (#789) | As QUEUE-12. |
| STOP-10 | pass (#789, #792) | Escape during the tool with a steer waiting: the steer became a Not sent row (not restored into the composer), the call settled 0.78 s later and the turn ended "Request aborted". |
| STOP-11 | pass (#789 changed) | Stop with a steer and /compact queued: toast "Stopped · 1 queued message not sent, kept above the composer; cancelled queued /compact"; the steer became a Not sent row and /compact did not run. Escape uses the same stop(). |
| STOP-12 | pass (#789 changed) | As TURN-18: compaction can now be stopped. |
| STOP-13 | pass (#789 changed) | Window shrunk to 49k on a 30.8k tab; a 400-row read triggered compaction between model responses; Stop during "Compacting context…" cancelled it: toast "Compaction stopped", turn ended "Request aborted" (no longer waits for the summary). |
| STOP-14 | pass (#795 changed) | "Stop and close" closed Chat 4 with "Closed Chat 4 · Undo"; Undo reopened it as "Chat 4" (its old number) with the partial answer and the aborted ending. |
| STOP-15 | fail | desktop (macOS), `8099cd6`: unchanged from `main`. During a streaming turn, Escape with focus in the grid stayed with Excel and the turn kept streaming; Escape with focus in the composer stopped it. |
| THINK-01 | pass | Popover on GPT-6 Luna: Off, Min, Low, Medium, High (ticked), Extra high, Max, each with a hint. |
| THINK-02 | pass | From High: Extra high, Max, Off, Min, Low, Medium, High, Extra high, each with "Thinking: {level} (next turn)". GPT-6 Luna offers Off, so it wraps to Off. |
| THINK-03 | pass (#786) | Finished block "Thought for 1s" when seen live. |
| THINK-04 | pass (#786) | "Thinking…" only while the block streamed (25 ms), then "Thought for 1s"; no "Thinking…" through the tools. |
| THINK-05 | pass (#793) | In Confirm mode (set from the new mode popover), Shift+Tab in the composer left the badge "Confirm". |
| THINK-06 | pass (#793 changed) | As TASKPANE-07: Shift+Tab cycles thinking only from the composer and otherwise moves focus normally. |
| THINK-07 | pass | As STOP-04. |
| THINK-08 | pass | Shift+Tab to "Extra high" during the first model response of a two-tool turn: all three requests of that turn went out at "high"; the next turn's request at "xhigh". |
| THINK-09 | pass | Chat 1 on "Low", "+": the new tab "Chat 9" started at "High" (GPT-6 Luna). |
| THINK-10 | pass (#786 changed) | The thinking header is now a focusable button: Tab through Chat 4's transcript reached "Thought for 2s" (aria-expanded=false), and Enter expanded it. |
| CTX-01 | pass | Meter "4%/272k" on GPT-6 Luna ("{pct}% / {window}"). |
| CTX-02 | pass | As TURN-10. |
| CTX-03 | pass (#789) | "Compacting context…", toast "Compacting to free up context", Enter queued a prompt; Stop button instead of send (#789). The "Compacting…" toast then stayed ~50 s after compaction ended (new bug, #793 stacked toasts). |
| CTX-04 | pass | "Summarized 1 message" toast; top of the transcript "Show earlier messages · 1 chat message" and "Summarized 122 tokens" cards (#798 kept most of the tail as compacted previews). |
| CTX-05 | pass (#789) | /compact queued behind a read turn in Chat 3, switched to empty Chat 4 before the turn ended: Chat 3 was compacted ("Summarized 8 messages", new summary); Chat 4 unchanged (1 message). The toasts still show on whichever tab is active. |
| CTX-06 | pass | "Too few messages to compact" shown for /compact in an empty tab (see CTX-08 run). |
| CTX-07 | pass (#789 changed) | As TURN-18. |
| CTX-08 | pass (#789 changed) | As SEND-11: "Compacting context…", Stop button (compaction is stoppable now), toast "Compacting to free up context", message visible as a Queued row. |
| CTX-09 | pass | Overflow injected ("…exceeds the context window…") on a ~50k-token tab: "Compacting context…", "Summarized 29 messages", the request was retried and answered "OVERFLOW-OK"; no banner left. (On a 30k tab where /compact itself says "Nothing to compact", overflow ends with the banner "Context window exceeded. Run /compact…" and no retry.) |
| CTX-10 | pass (#789) | /compact ("Summarized 4 messages"), then a taskpane reload at once: Chat 3 came back compacted (same 18 messages, summary first, meter 2%). |
| CTX-13 | pass (#789) | Meter 15% → 2% within 12 ms of "Summarized", no tab switch needed. |
| CTX-14 | pass (#789) | "Summarized 1 message" (singular). |
| CTX-11 | pass (#789 changed) | Colours now follow the compaction budget (window shrunk in memory, 11,668 tokens used): 35k window (33%) no colour; 33.8k (35%) yellow (above 75% of the 16.5k soft budget); 18.7k (62%) red; 14.6k (80%) red; 9.3k (125%) red with "Context is full — the next message will fail". |
| CTX-12 | pass | Tooltip "11,668 / 33,800 tokens · Context 35% full. Consider using /compact…"; red: "Context 62% full — responses may become less reliable. Use /compact…". Read from the DOM, not hovered. |
| AUTOCTX-01 | pass | Selected D7, typed 555 into B4 in the Excel grid: answer "Selected: Sheet1!D7. You changed Sheet1!B4 from 120 to 555."; no tool card. |
| AUTOCTX-02 | pass (#792 changed) | Request carried stored notes "Workbook structure (initial)", "Files workspace (initial)" and "Edits made in Excel… By the user: Sheet1: B4 (120 → 555)", plus a final selection note: address and size, a sample of the selection with up to 5 rows around it, "Formulas in the sample", "Errors in the sample: D10=#VALUE!". |
| AUTOCTX-03 | pass (#792 changed) | Two-tool turn, 3 requests: each ended with a freshly read selection note; the workbook/files/edit notes are stored once in the transcript and not rebuilt. |
| AUTOCTX-04 | pass (#792) | After Pi wrote Sheet1!G2, the next request had no edit note: Pi's own writes are no longer reported as user changes. |
| AUTOCTX-05 | pass (#792) | The blueprint was sent once as "Workbook structure (initial)" and not re-sent as "context missing from history". |
| AUTOCTX-06 | pass (#792 changed) | A grid edit to C6 is now reported to every tab: Chat 11's next request listed "By the user: Sheet1: B4 (120 → 555), C6 (220 → 777)" and "By Pi in another chat tab: G2"; Chat 13's next request also listed C6. |
| AUTOCTX-07 | pass (#792 changed) | A sheet added outside Pi (HandSheet3) produced a new stored note "Workbook structure (structure changed)" in the next request. |
| AUTOCTX-08 | pass (#792 changed) | After a file was dropped, the next request stored a new note "Files workspace (files changed)" (now counting files without a workbook tag); the request after had no new files note. A new tab gets "Files workspace (initial)" (Chat 13). |
| AUTOCTX-09 | pass (#792 changed) | A multi-area selection (A1:A2, C3:C4) gives a selection note "Unavailable: The current selection is invalid for this operation." instead of leaving it out silently, as #792 intends. |

## Regressions and new bugs

Each was sent to its owning PR's thread as it was found. Line numbers are at `e48d83c`.

1. **SEND-03 regression: a steer turns following back on (#789). Open on `e48d83c`; fixed in #789 @ 4354892, passes on `8099cd6`.** Repro: scroll the transcript up 400 px while Pi is responding, then send a steer with Enter. The transcript jumps to the bottom. On main it stayed put. Cause: `_onSend` sets `_autoScroll = true` for every submit (src/ui/pi-sidebar.ts:664). The owner confirmed this was not deliberate.
2. **COMPOSE-11 still fails (#789). Open on `e48d83c`; fixed in #789 @ 4354892, passes on `8099cd6`.** Repro: transcript at the bottom; in the composer type "g0", Shift+Enter, "g1", Shift+Enter, "g2", Shift+Enter. Following turns off after the third line (gap 59–98 px). Cause: src/ui/pi-input.ts:182 sets `height: auto` while measuring, Chrome clamps scrollTop, and the scroll listener (src/ui/pi-sidebar.ts:644-660) takes it for the user scrolling up.
3. **Tab scroll arrows snap back (#793). Open on `e48d83c`; fixed on `8099cd6`** ("‹" reaches scrollLeft 0 and stays there after 3 s, and Chat 1–3 can be clicked). Repro: open 6 tabs in a 349 px pane with the last one active, then click "‹". `scrollBy(-140)` runs, then `updated()` scrolls the active tab back into view (src/ui/pi-sidebar.ts:400, on any `sessionTabs` change). scrollLeft goes 199 → 175 on every click, so earlier tabs can't be reached with the pointer. Clicks on the half-hidden tabs are also intercepted.
4. **/files: the first Escape drops focus to `<body>` (#793). Open on `e48d83c`, minor; fixed on `8099cd6`.** Repro: `/files`, Enter (focus in the filter field), Escape. The overlay stays open with focus outside it; a second Escape closes it. The Escape handler blurs the text field (src/taskpane/keyboard-shortcuts.ts:202) before the overlay sees the key.
5. **"Compacting to free up context" toast lingers (#793 stacked toasts). Fixed on `e48d83c` for a finished compaction; after Stop, fixed on `8099cd6`** ("Compaction stopped" replaces it and is gone about 4 s later). Repro: `/compact`, then Escape. "Compaction stopped" appears, and the 60 s "Compacting…" toast stays beside it, because src/commands/builtins/export.ts:851 does not pass `id: COMPACT_TOAST_ID`.
6. **Toasts covered the Not sent rows (#793). Fixed on `e48d83c`.** On a2b3f2d `--pi-dock-height` was never set (`_observeDock` ran in `firstUpdated`, before an agent was bound). The Stop toast sat over the rows' Edit/× buttons and paused when hovered, so × could not be clicked. On `e48d83c` the toast sits above the rows (418–468 vs 484–517 px) and × is hit directly.
7. **"+" New tab hidden under the gear with three tabs (#793). Fixed on `e48d83c`**: "+" sits outside the scroller and is hit at its centre.
8. **Proxy banner missing when the proxy is enabled while already down (#793). Fixed on `e48d83c`**: turning the Proxy switch on and off now shows and hides the banner.
9. **The model answered the selection note instead of the request (#792; the combination made it visible). Suspected, flaky. Fixed on `e48d83c`.** On a2b3f2d, 1 of 5 fresh GPT-6 Luna tabs answered the three-read prompt with "What would you like me to do with the workbook?". The trailing user-role selection note came after the tool results. On `e48d83c` the note says it is background and restates the request; 5 of 5 tabs answered DONE.
10. **Hover × on a narrow inactive tab covers the title, so a click on the title closes the tab (#793). New on `8099cd6`; fixed on `cea7588`** (#793 @ c88b997 restores main's in-flow ×). Repro: open about seven tabs so the inactive ones are 55–60 px wide. Move the mouse onto "Chat 5", then click 4 px right of its centre, on the title text. The tab closes with no prompt. While the tab is hovered, `elementFromPoint` at the centre returns `__title`, at centre+2 px `__close`, which spans offsets 29–51 of the 55 px tab. Cause: src/ui/theme/components/tabs.css:255-265 places the inactive tab's × with `position: absolute; right: 0` over the 37 px title, and `:hover` turns on its pointer events. Earlier in this run the same thing closed Chat 1 by accident. Sent to #793.

Resolved: **SEND-13 (#789)**, fixed by #789 @ a66868d (passes on `8099cd6`). Before that, when the proxy dies mid-stream, the retry path drops the failed response (src/taskpane/action-queue.ts:247), so the partial answer is lost when every retry fails. On main it was kept with `stopReason: error`. Sent to #789's owner to confirm whether this is intended; recorded as a pass meanwhile.

## Rechecked on `e48d83c`

- Compaction of a session whose only user message is archived (`/compact` on a 16% tab: "Summarized 1 message", no user message left), then a reload: the summary and kept messages came back unchanged, meter 2%.
- TABS-22 shape: Chat 10 asked to write H1 with `Excel.run` delayed 8 s. Its × opened "Stop and close this tab?"; the tab took the workbook lock while the dialog was open; then "Stop and close". The toast "Stopping — closing after Pi's current change finishes." showed, and the tab closed about 5.8 s later with "Closed Chat 10 · Undo". The stopped call reported an unknown outcome, H1 stayed empty, and the lock was free: the next write, H2 = 2, went through.
- Toasts against Not sent rows, "+", proxy banner, compaction toasts and the selection note: see the list above.

## Rechecked on `8099cd6`

Dev server restarted from `8099cd6` (#789 @ a66868d, #793 @ aff33d6, #792 @ d31bf82, #794 @ 515072d) and confirmed serving it.

- COMPOSE-11, SEND-03, SEND-13: pass (see their rows).
- Compaction toast: during `/compact`, "Compacting to free up context" shows; Escape replaces it with "Compaction stopped", which is gone about 4 s later. No 60 s toast remains.
- Tab strip: "‹" scrolls to the first tabs and stays there; a pointer click on Chat 2 switched to it. Shift+F10 opens the tab menu, and Tab closes it with focus on the tab; Escape while Pi is responding closes the menu and the turn continues. New bug 10 (hover × over the title).
- /files: TASKPANE-05/06 pass; Escape and Tab stay inside the overlay.
- Slash menu: `/` opens `#pi-command-menu` (role listbox, 26 commands); the textarea has `aria-controls="pi-command-menu"` and `aria-activedescendant`. A pointer press on the transcript closes it and keeps the draft. After switching tabs, typing `/sh` in the new tab's composer filters to /share-session and /shortcuts, and a pointer click on an item runs it. While Pi is responding, Escape with `/shortcuts` typed keeps the draft and the turn; Enter gives "Wait for this tab to finish before running /shortcuts" and keeps the draft.
- Stop during tool calls (#792 grace, `Excel.run` wrapped): a delayed read settled 2.6 s after Stop with "Stopped by the user before Excel answered.". write_cells with Excel.run delayed 4 s before running: "Stopped before writing — nothing changed in Sheet1!H4", cell empty, turn ended 5.5 s after Stop. write_cells where Excel wrote at once but answered 15 s late: write_cells reports Stop itself (`reportsStop`, so no 10 s cut-off), and the turn ended 14.8 s after Stop with "Written to Sheet1!H15 … Stop was pressed while this ran; the change above was already applied and stands." and a backup id; the cell holds 15000. That matches #792 and DECISIONS.md (bounded by the 60 s `excelRun` timeout).

## Rechecked on `cea7588`

Dev server restarted from `cea7588` (#793 @ c88b997, #792 @ c8593dc, #791 @ e1c3f29); the served tabs.css has the inactive × in the flow (`position: static`).

- Narrow inactive tab title click (bug 10): fixed. Tabs are 80–85 px with the × at offset 54–76, after the title (9–46), and never over it. Real-pointer clicks on the right edge of the title of Chat 11, Chat 10 and Chat 3, and 4 px right of Chat 2's centre, each switched to that tab; all 6 tabs stayed open. The × still closes: 3 of 3 times a new inactive tab closed on a real click on its ×, with "Closed Chat N · Undo".
- Slash menu row click while a toast shows (#793 P-165): pass. With a toast already up when `/` opened the menu, the toast moved above it (menu top 276 px, toast bottom 268 px, `--pi-command-menu-offset` 346px). The row under the pointer was the row itself, and a real click on /rules, after moving over the rows above it, opened Rules. With a toast shown while the menu was open, the toasts stacked above it (196–268 px), and a click on /tools opened Connections. With `/sh` (a 78 px menu) and two toasts, a click on /shortcuts opened Keyboard shortcuts. After the menu closed the offset was removed, and the toasts returned above the composer.

## Not run here (desktop only)

- TASKPANE-14: the composer has focus when Excel desktop opens the taskpane.
- TURN-22: whether a write already handed to Excel completes after Stop.
- COMPOSE-13: IME composition Enter on macOS (#789's `keyCode 229` check).
- STOP-15: Escape reaching the taskpane in every focus state on Excel desktop.
