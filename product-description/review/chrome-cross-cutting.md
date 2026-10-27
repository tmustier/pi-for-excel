# Review: chrome and cross-cutting behavior

This review asks whether the behavior described is the best behavior for users. It does not ask whether the description is accurate. It covers the status bar, menus, overlays, toasts and banners, message rendering, keyboard shortcuts, language, persistence, provider errors and host differences (`ui/*`, `cross-cutting/*`).

The target is `main` as documented plus the fixes in flight, read from the PR branches: #786 `fix-transcript` at `30b389b`, #789 `fix-turn-composer` at `478823c`, #793 `ux-system` at `49999ef`, and the [taskpane UX system](https://t-03h4e6pcg0q4sx8jabo84ftco-p22710.onamp.dev/taskpane-ux-system.html) design. Defects already in [bug-triage.md](../bug-triage.md) are not raised again. They are cited only where a proposal changes the decision on one.

The users in mind are analysts and finance professionals working in a narrow pane (about 340 px) on Excel desktop or Excel on the web, including keyboard-only and screen-reader users.

## Summary

The fix PRs move the chrome to a sound model. Each tab owns its draft, error and busy state (#789). There is one contract for overlays, menus and notices (#793). The transcript is drawn once and has a real details mode (#786). What remains are mostly gaps the PRs carry forward from `main`:

1. **A failed turn is still a dead end.** A failed turn gets a top banner, plus an "Error:" line in the transcript with the provider's raw text, and neither offers a way forward. Retry is short and fixed: 14 s in total, ignoring Retry-After.
2. **The context meter nags about something the product already handles.** It tells users to run `/compact` and warns that "the next message will fail", while auto-compaction makes both untrue.
3. **Analysts lack two things that comparable products treat as basic:** clickable cell references in Pi's own answers, and copying an answer or a table back into Excel.
4. **Keyboard and screen-reader users still lose out.**
   - Nothing announces when a turn finishes, fails, or waits for approval.
   - Shift+Tab is still taken from standard focus navigation inside the composer.
   - On the web, the tab shortcuts are hidden and nothing replaces them.
   - Ctrl/Cmd+Z reopens a closed chat when Excel users expect undo.

## Proposals, ranked

Each proposal names its type: **fix** (clearly better, low debate) or **product call** (a trade-off for Thomas to decide).

### CC-01: One inline error with "Retry", in place of the run-error banner

- **Today / planned:** With #789 and #793, a failed turn shows a red, dismissible banner in the notice area under the tabs, on that tab only. The failed response also ends with "Error:" and the provider's reason. Neither offers an action. The user retypes "try again". A 401 shows the provider's raw JSON in both places ([provider errors and offline](../cross-cutting/provider-errors-and-offline.md#the-request-fails), [message rendering](../ui/message-rendering.md#errors-and-endings)). `showErrorBanner` in `ux-system` passes only a message, with no actions (`src/taskpane/init.ts:213`).
- **Proposed:**
  - Show the failure once, at the failed response. Give it a one-line plain-language reason and a **Retry** button that continues the turn without retyping.
  - When the cause is known, add one fix action: "Update API key" for 401/403, "Open proxy settings" when the proxy is unreachable, "Choose another model" for 404 or an outage.
  - Put the provider's raw text behind "Details".
  - Keep the notice area for states that remain true (offline, proxy down, missing login). A finished failed turn is not a standing state.
- **Why:**
  - #793's own taxonomy (D2: "toast for 'done', banner for 'still true'") puts a failed run inline, not in a banner.
  - In a 340 px pane the banner and the inline line repeat the same text, and the banner pushes the transcript down.
  - ChatGPT, Claude.ai and Cursor all put a Retry or "Try again" button on the failed message.
  - Raw JSON such as `OpenAI API error (401): {…}` means nothing to a finance user.
- **Cost / risk:** Medium. Retry is `agent.continue()` after dropping the failed response, which #789's retry loop already does. The error mapping for 401/403/404, offline and proxy already partly exists in `turn-error.ts`. The banner code in #793 can stay for the standing states.
- **Touches:** #789 (`turn-error.ts`), #793 (banner use).
- **Type:** product call (dropping the run-error banner); fix (Retry button, plain reason).

### CC-02: Make automatic retry match how providers actually fail

- **Planned (#789):** Retry up to 3 times at 2 s, 4 s and 8 s, which is 14 s in all. Failures are classed by pi-ai's `isRetryableAssistantError`, and the provider's Retry-After is ignored. A mid-stream failure is also retried: the partial answer is dropped and the turn continues. ChatGPT reports every 429 as "You have hit your ChatGPT usage limit." That text matches neither of pi-ai's patterns, so a transient ChatGPT 429 is still never retried (P-77 is only half fixed).
- **Proposed:**
  - Honour Retry-After when it is under about 60 s.
  - Otherwise back off over about a minute, for example 6 attempts. Overload episodes (529, 503) usually last longer than 14 s.
  - Keep failing fast on quota and billing errors.
  - Tell a ChatGPT usage-limit 429 apart from a short rate limit, using its error code rather than the rewritten message.
  - Do not discard a partial answer the user has already read. Keep it, mark it interrupted, and continue.
  - When retries run out, fall through to CC-01's Retry button.
- **Why:**
  - Claude Code retries 10 times by default. It retries only failures "that arrive before any of Claude's response has streamed", and it fails at once on spend-limit 429s ([Claude Code error reference](https://code.claude.com/docs/en/errors)). That is the same split proposed here.
  - 14 s turns a routine overload into a failed turn. Text that vanishes mid-read looks like a bug.
- **Cost / risk:** Low. It changes the constants and the classification in `action-queue.ts:217`. The countdown UI ("Retrying in Ns (attempt n of N)…") already exists.
- **Touches:** #789.
- **Type:** fix (Retry-After, budget, ChatGPT 429); product call (how to handle the partial answer).

### CC-03: A quiet context meter when auto-compaction is on

- **Planned (#789):** Colours follow the compaction budget: on a 200k model, yellow above 120k and red above 160k, with compaction at 170k. The copy is unchanged from `main`: "Consider using /compact to free space or /new to start fresh.", "Context {pct}% full — responses may become less reliable." and "Context is full — the next message will fail." The popover's buttons are labelled "/compact" and "/new" ([the status bar](../ui/the-status-bar.md#the-context-meter), [context and compaction](../turn/context-and-compaction.md#the-meter)).
- **Proposed:**
  - With auto-compaction on (the default, with no setting to turn it off), drop the advice and the "will fail" warning.
  - Near the threshold, say what will happen: "Pi will summarize older messages automatically soon."
  - Label the popover actions in plain words, "Summarize now" and "New chat", with the command names as secondary text.
  - Keep the "full" warning only where it is true: auto-compaction off, or a second overflow.
- **Why:**
  - The status bar is the most visible chrome in the pane. Turning it red and telling users to act on something the product does by itself is noise that teaches people to ignore red.
  - "The next message will fail" is false while auto-compaction runs before sending (confirmed in the status-bar pass: 100% was never reached because compaction ran first).
  - Slash names are jargon to the target user.
- **Cost / risk:** Low: copy and one condition in `status-bar.ts`, plus zh-CN strings.
- **Touches:** #789 (meter colours), #793 (status bar).
- **Type:** fix.

### CC-04: Make cell references in Pi's answers clickable

- **Today / planned:** Only addresses inside tool cards are links. "see Sheet1!B12" in Pi's text is plain text ([message rendering](../ui/message-rendering.md#cell-links)). #786 does not change this.
- **Proposed:**
  - Turn sheet-qualified references in assistant text into the same cell links tool cards use, for example `Sheet1!B12`, `'Q3 Plan'!C4:F9`, or a bare `B12` when the answer is about the active sheet.
  - Link only references that parse as valid A1 notation and name an existing sheet.
  - Leave code spans and code blocks alone.
- **Why:**
  - The most common answer to "explain this model" points at cells. Analysts check each claim against the grid, and here they have to find each cell by hand.
  - Claude for Excel advertises this as a core feature: "Answers include cell-level citations you can click to navigate to the referenced cell" ([Claude Help Center](https://support.claude.com/en/articles/12650343-use-claude-for-excel)).
- **Cost / risk:** Medium. It needs a markdown post-pass over text nodes, plus a sheet-name check that has to stay cheap while text streams; resolving after the response completes is enough. False positives (a word like "Q3" or "FY24") are avoided by requiring A1 shape and a known sheet. Cell-link click handling and error toasts already exist (#786).
- **Touches:** #786 (`markdown-block.ts`, `cell-link.ts`).
- **Type:** product call (scope), then fix.

### CC-05: Copy an answer, and copy tables so they paste into cells

- **Today / planned:**
  - The only per-item copy is the code-block "Copy" button.
  - Copying a whole answer is `/copy`, which takes the last answer only and is refused while busy (P-113).
  - Markdown tables in answers have no copy at all.
  - On Excel on the web, the code block reports "Copied!" even though the clipboard API is blocked there. Whether anything was copied is unverified ([message rendering](../ui/message-rendering.md#code-blocks)).
  - #795 gives `/copy` an execCommand fallback and then a manual-copy dialog, but the code block keeps its own path.
- **Proposed:**
  - Give each assistant answer a Copy action, shown on hover and focus.
  - Give each rendered table a "Copy table" action that writes tab-separated text, so pasting into Excel fills cells.
  - Route every copy through one helper (#795's): clipboard API, then execCommand, then the manual-copy dialog. Show "Copied!" only on confirmed success.
- **Why:**
  - The finance workflow ends outside the chat. Summaries go into email or a memo, and tables go back into a sheet.
  - Per-message copy is standard in ChatGPT, Claude.ai and Copilot chat.
  - A tab-separated table is what Excel expects on paste, so a table becomes a range with no reformatting.
  - A false "Copied!" is worse than none.
- **Cost / risk:** Low to medium. It is UI only and reuses #795's helper.
- **Touches:** #786 (message components), #795 (copy helper).
- **Type:** fix (one honest copy path); product call (per-answer and table copy).

### CC-06: Announce turn progress to screen readers

- **Today / planned:** Toasts are live regions (polite, errors as `alert`), and #793 makes error banners `role="alert"`. The streaming text, the working indicator, the end of a turn, the retry countdown, and an approval waiting in another tab are not announced. `aria-live` appears only on status popovers and setup cards on all three branches.
- **Proposed:** Add one polite status region that announces the following:
  - "Pi is working…" once
  - "Pi finished" or "Pi stopped" at the end of a turn, with the first sentence of the answer optional
  - "Retrying in N s"
  - "Waiting for your approval in Chat 2"

  Do not stream tokens into it.
- **Why:**
  - Turns take tens of seconds to minutes. A screen-reader user has no way to know when to come back except by polling the transcript.
  - WCAG 2.2 SC 4.1.3 Status Messages requires these to be programmatically announced without moving focus ([W3C Understanding 4.1.3](https://www.w3.org/WAI/WCAG22/Understanding/status-messages.html)).
- **Cost / risk:** Low. It is one region fed from the per-tab activity state #789 already has.
- **Touches:** #789 (activity state), #793 (notice primitives).
- **Type:** fix.

### CC-07: Take thinking-level cycling off Shift+Tab

- **Planned (#793, Q4):** Shift+Tab cycles the thinking level only when focus is in the composer and no overlay or menu is open. Everywhere else it moves focus backwards.
- **Proposed:**
  - Leave Shift+Tab as standard reverse focus navigation everywhere.
  - Move cycling to a chord the composer does not otherwise use, for example Ctrl/Cmd+Shift+. or Alt/Option+T. Check it against Excel and the browser first (see CC-08).
  - Keep the clickable thinking control as the main way to change the level.
- **Why:**
  - The composer is where keyboard users spend their time, so it is exactly where reverse navigation matters. From there, Shift+Tab is how to reach the Files button, a starter hint or the transcript.
  - #793 removes the worst case (focus jumping plus a silent level change on the web), but a keyboard user moving back from the composer still changes how slow and expensive the next turn is, and never moves.
  - The WAI-ARIA Authoring Practices treat Tab and Shift+Tab as the focus-movement keys a page should not repurpose ([APG: Developing a keyboard interface](https://www.w3.org/WAI/ARIA/apg/practices/keyboard-interface/)).
  - Claude Code uses Shift+Tab for mode cycling, but in a terminal that has no focus order.
- **Cost / risk:** Low in code. Users who learned Shift+Tab lose it, so show the new chord in the thinking tooltip and in the toast.
- **Touches:** #793.
- **Type:** product call (this reverses accepted default Q4).

### CC-08: One host-neutral set of tab shortcuts, and no Ctrl/Cmd+Z overload

- **Planned (#793):**
  - On Excel on the web, the shortcuts page hides Ctrl/Cmd+T, Ctrl/Cmd+W, Ctrl/Cmd+Shift+T and Ctrl+PageUp/PageDown, and tells users to use "+" and the tab menu.
  - On desktop the same chords work. Ctrl/Cmd+W in a non-empty composer closes the chat tab.
  - Ctrl/Cmd+Z reopens the last closed tab whenever the composer is empty or the undo toast is showing ([keyboard shortcuts](../ui/keyboard-shortcuts.md#tabs)).
- **Proposed:**
  - Pick tab chords that reach the taskpane on every host. Alt/Option+N, Alt/Option+W and Alt/Option+Shift+T are candidates, to be checked against Excel's Alt KeyTips on Windows desktop.
  - Use the same chords everywhere and list them the same way.
  - Keep the native chords on desktop as aliases only if they test cleanly.
  - Drop Ctrl/Cmd+Z as "reopen closed tab". The undo toast's button, Ctrl/Cmd+Shift+T (or its replacement) and `/reopen` remain.
- **Why:**
  - Today a web user has no keyboard path to a new, closed or reopened tab, and a desktop user's muscle memory does something different on the web. There, Ctrl+W closes the whole Excel browser tab.
  - Microsoft's add-in guidance says to "avoid registering keyboard shortcuts … already in use", and that browser shortcuts cannot be overridden on the web ([Custom keyboard shortcuts in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts)).
  - Excel users press Ctrl+Z to undo. In the taskpane, with the composer empty, it silently reopens a chat they closed, or does nothing, when they meant to undo Pi's last change.
- **Cost / risk:** Low in code. The real work is a key-collision check on three hosts.
- **Touches:** #793 (shortcuts page, keyboard model).
- **Type:** product call.

### CC-09: Show on the tab when a background tab needs attention

- **Planned (#789, #793):** A background tab's error is kept and shown when the user switches to it. The busy dot (made visible by #793) goes off when the turn ends. Nothing on the tab says whether it finished, failed or is waiting. An approval from a background tab opens over the active tab and names the tab (#794).
- **Proposed:** Give the tab label two small states beside the busy dot.
  - **Failed:** a red dot, kept until the user views the tab.
  - **Finished, not yet seen:** a neutral dot, kept until viewed.

  Announce both through CC-06.
- **Why:**
  - Per-tab errors fix the wrong-tab banner (P-102), but they also make a background failure invisible until the user happens to switch to that tab.
  - The reason to run two tabs is to start one job and do another, so "is the other one done?" is the first question.
  - Cursor and Claude desktop mark background chats that have new output.
- **Cost / risk:** Low. The state already exists per tab in #789 (`tab-state.ts`), and #793 owns the tab visuals.
- **Touches:** #789, #793.
- **Type:** fix.

### CC-10: Follow the Office theme by default

- **Today / planned:** The taskpane is always light. Dark mode sits behind `/experimental` and, on the web, never turns on, because Office reports a white background (P-135). `ux-system` still gates it on the experiment flag (`src/ui/theme-mode.ts:50`).
- **Proposed:**
  - Graduate dark mode and follow Office's own theme (`Office.context.officeTheme`) by default, with a Light / Dark / Follow Office choice in Settings.
  - On the web, where Office reports no usable theme, stay light.
- **Why:**
  - A bright white panel beside a dark Excel grid is the first thing a dark-theme user sees. Copilot in Excel follows the Office theme.
  - The palette and the detection already exist, so the cost is mostly testing every surface #793 just unified.
- **Cost / risk:** Medium: a visual pass over every surface in both themes. #793's unified tokens make this the cheapest time to do it.
- **Touches:** #793.
- **Type:** product call.

## Right as planned; do not change

- **Markdown safety rules.** Raw HTML is shown as text, images are never loaded, only http/https/mailto/tel links are live, and `$` is never math ([message rendering](../ui/message-rendering.md#ends-at-once-what-is-not-rendered)). Spreadsheets are full of `$`, and a taskpane that can write to the workbook must not load remote content from model output.
- **Stop is one action that gives back what the user wrote (#789).** Queued text returns to the draft, queued commands are cancelled and named, and nothing undelivered is ever sent unasked. This is the right contract and matches how Claude Code and Cursor treat interruption.
- **Per-tab draft, error and activity (#789).** Tabs are separate jobs, and anything held per taskpane leaked between them (P-48, P-102).
- **One overlay contract (#793).** Focus moves in and is trapped, Escape affects only the topmost overlay, focus returns to the opener, and shortcuts are off under a modal. This is the standard dialog pattern, and it removes a long tail of "typing went to the composer behind the dialog" bugs (P-22).
- **Notification taxonomy and wording (#793 D2/D5).** Toasts confirm actions and pause on hover, errors persist longer and can be dismissed, and changes made from a control show in place with no toast. This keeps a narrow pane quiet. CC-01 applies the same taxonomy more strictly.
- **Auto/Confirm as a popover, not a toggle (#793 Q3).** The control changes whether Pi asks before writing to the workbook, which deserves a deliberate choice with one line of explanation. A one-click toggle in the status bar was too easy to flip by accident.
- **The proxy banner only when the proxy is enabled (#793 Q1).** It stops warning API-key users about a helper they never use. Failures caused by a missing proxy now get specific messages (#789).
- **Details mode starts collapsed and is remembered (#786).** It matches how cards start, a hand toggle wins for one item, and Ctrl/Cmd+O matches Claude Code's transcript toggle.
- **Agent-facing text is never translated.** It keeps the prompt prefix identical across UI languages, so caching works and the model's behavior does not change with a display setting.
- **Persistence is silent when it works and loud when it fails (#789's save-failure toast).** Analysts should never think about saving chats. They need to know only when something was lost.

## How this was reviewed

- I read the `ui/` and `cross-cutting/` documents, goal.md's "Fixes in flight", bug-triage.md, the PR descriptions of #786, #789 and #793, and the UX system design (decisions D1–D9, Q1–Q4).
- I checked claims about planned behavior against the PR branches at the SHAs above: `turn-error.ts`, `action-queue.ts`, `toast.ts`, `banner.ts`, `status-bar.ts`, `shortcuts-page.ts`, `theme-mode.ts` and en.json, and pi-ai's `utils/retry.js` for the retry classification.
- Running-product observations are the ones already confirmed in the documents. No new session was run for this review.
