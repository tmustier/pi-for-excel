# Review: the core interaction loop

Scope: sending, steering and the queue, Stop, thinking, compaction, tabs, resume and reopen, the command menu and the keyboard model ([the taskpane](../foundations/the-taskpane.md), [turns](../foundations/turns.md), [sessions and tabs](../foundations/sessions-and-tabs.md), `turn/`, `sessions/`, `commands/`). The question is whether the behavior is the **best** behavior for users, not whether it is described correctly. The baseline is `main` as documented, plus the fixes in flight ([#786](https://github.com/tmustier/pi-for-excel/pull/786)–[#795](https://github.com/tmustier/pi-for-excel/pull/795) and the [taskpane UX system](https://t-03h4e6pcg0q4sx8jabo84ftco-p22710.onamp.dev/taskpane-ux-system.html)). Defects already in [bug-triage.md](../bug-triage.md) are not raised again.

The users: analysts and finance professionals in Excel desktop or on the web, often in a narrow pane, sometimes co-authoring, trusting an agent with a workbook they are accountable for.

## Summary

Once the fix PRs land, the loop is coherent. Each tab owns its own state, Stop is one action, nothing is sent behind the user's back, and overlays own their keys. What remains is mostly **trust and simplicity**, not correctness:

1. After a turn, the user cannot tell what Pi changed or undo the turn as a whole. That is the question an analyst asks after every Stop or failure.
2. Talking to Pi while it works needs two invisible keyboard modes (steer and follow-up). For a mouse user, the button where Send used to be is now Stop.
3. Several defaults and shortcuts work against Excel habits:
   - Ctrl/Cmd+Z can reopen a tab.
   - Shift+Tab changes the thinking level instead of moving focus.
   - The working line shows jokes instead of what Pi is doing.
   - Tabs stay "Chat N" forever.
4. The command and session surface has more paths than users need: `/resume-here`, the replace target, the fork setting, aliases, and a "coming soon" command.

There are twelve proposals, ranked by user impact. Five need a product call from Thomas; seven are fixes.

## Proposals

| ID | Proposal | Type | Touches |
| --- | --- | --- | --- |
| [CL-01](#cl-01-show-what-a-turn-changed-and-let-the-user-undo-the-turn) | Show what a turn changed, and offer "Undo this turn" | product call | #794 |
| [CL-02](#cl-02-ctrlcmdz-must-not-reopen-tabs-unless-the-undo-toast-is-showing) | Ctrl/Cmd+Z reopens a tab only while the Undo toast shows | fix | #795 |
| [CL-03](#cl-03-one-visible-way-to-talk-to-pi-while-it-works) | One visible way to talk to Pi while it works; Send next to Stop | product call | #789 |
| [CL-04](#cl-04-make-queue-rows-readable-and-individually-editable) | Queue rows readable and individually editable or removable | fix | #789, #793 |
| [CL-05](#cl-05-the-working-line-says-what-pi-is-doing-including-stopping) | The working line says what Pi is doing, including "Stopping…" | product call | #789, #793 |
| [CL-06](#cl-06-a-background-tab-that-finishes-fails-or-needs-approval-is-marked) | Mark a background tab that finished, failed or needs approval | fix | #789, #793, #794 |
| [CL-07](#cl-07-retry-a-failed-turn-and-recall-the-last-message) | Retry a failed turn; Up in an empty composer recalls the last message | fix | #789 |
| [CL-08](#cl-08-tabs-are-titled-by-their-conversation) | Tabs titled by their conversation, not "Chat N" | product call | #795 |
| [CL-09](#cl-09-shifttab-moves-focus-thinking-gets-its-own-key) | Give Shift+Tab back to focus; thinking gets its own key | product call | #793 |
| [CL-10](#cl-10-resume-always-opens-or-focuses-a-tab-drop-replace-current-tab-and-resume-here) | Resume always opens or focuses a tab; drop "Replace current tab" and `/resume-here` | fix | #795 |
| [CL-11](#cl-11-model-changes-while-busy-apply-from-the-next-turn-like-thinking-drop-the-fork-setting) | Model changes while busy apply from the next turn; drop the fork setting; new tabs inherit thinking | fix | #795 |
| [CL-12](#cl-12-prune-and-rank-the-command-menu-refuse-only-what-conflicts) | Prune and rank the command menu; refuse a command while busy only when it conflicts | fix | #789, #795 |

---

### CL-01: Show what a turn changed, and let the user undo the turn

- **Today / planned:** Every mutation has its own [automatic backup](../workbook/backups-and-revert.md). After a turn, including a stopped or failed one, "changes Pi made … stay made" ([turns › Finishes](../foundations/turns.md#finishes); [stopping › After stopping](../turn/stopping.md#after-stopping)). To find out what happened, the user reads the tool cards. To undo it, they run `/revert` once per backup, newest first, or open the Backups dialog. #794 makes backups reliable but keeps them per mutation.
- **Proposed:** When a turn that changed the workbook ends, show one line under it: "Changed Sheet1!B2:B40, Summary!A1 (3 changes) · Review · Undo this turn". "Undo this turn" restores that turn's backups in reverse order and leaves an ordinary "Restore" backup, so it can itself be undone. Show the line on every ending, and make it most prominent after Stop or a failure.
- **Why:** After Stop, the first question is "what did it already do to my workbook?". Today the answer is spread across collapsed tool cards (3 or more start collapsed under #786), and undoing it means counting `/revert`s. Excel's own Undo may or may not cover Pi's writes, depending on the host (it does on the web and Mac, per goal.md). Comparable products make undo per turn, not per write:
  - Cursor: "Restore Checkpoint … to roll back all changes Agent made after that point" ([Cursor help](https://cursor.com/help/ai-features/agent)).
  - Claude Code: checkpoints taken before each prompt, rewound with Esc Esc or `/rewind` ([checkpointing](https://code.claude.com/docs/en/checkpointing)).
  - Copilot in Excel only points users to Undo and version history ([FAQ](https://support.microsoft.com/en-us/office/frequently-asked-questions-about-agent-mode-in-excel-1cfd906d-40b4-46be-8e2d-65b893e28a02)), which is the weaker model this proposal goes beyond.
- **Cost / risk:** Medium. Backups already carry the tool call; they also need a turn ID. Restoring in reverse order across ranges the user edited afterwards needs the same overwrite rules as `/revert`. It touches the Backups UI, which #794 owns.

### CL-02: Ctrl/Cmd+Z must not reopen tabs unless the Undo toast is showing

- **Today / planned:** Ctrl/Cmd+Z reopens the most recently closed tab whenever one exists and focus is not in a text field, or the text field is empty ([keyboard shortcuts](../ui/keyboard-shortcuts.md); [reopening closed tabs › Choosing what to reopen](../sessions/reopening-closed-tabs.md#choosing-what-to-reopen)). On `fix-sessions-models` the guard is unchanged (`shouldHandleUndoCloseTabShortcut` returns true for an empty composer), and #795 now **persists** recently closed per workbook.
- **Proposed:** Ctrl/Cmd+Z reopens a tab only while the "Closed {title} · Undo" toast is showing. At any other time it does nothing in the taskpane.
- **Why:** Ctrl+Z is the most practised key in Excel. A user who just sent a message has an empty composer with focus. If they press Ctrl+Z to undo a cell Pi wrote, an old conversation pops open as a new tab. With #795 that tab can be one closed days ago. The undo they wanted does not happen, and the result looks unrelated to what they did. Ctrl/Cmd+Shift+T, `/reopen` and the picker still cover deliberate reopening.
- **Cost / risk:** Trivial: one condition. No loss of capability.

### CL-03: One visible way to talk to Pi while it works

- **Today / planned:** While Pi responds, there are two ways to send:
  - Enter **steers**: delivered after the current tool batch.
  - Alt/Option+Enter sends a **follow-up**: delivered only when Pi would otherwise finish.

  Both are keyboard-only. The send button turns into Stop, so "steering and following up are keyboard-only" ([steering and the queue › What Enter means while busy](../turn/steering-and-the-queue.md#what-enter-means-while-busy)). The copy disagrees with itself:
  - placeholder: "Redirect reply (↵) · New question (⌥↵)", with ⌥ on Windows too;
  - Shortcuts page: "Interrupt and redirect", although a steer never interrupts;
  - working-line hint: "Press Enter to redirect Pi".

  #789 keeps both modes.
- **Proposed:**
  - **One send mode while busy.** Enter (and a visible Send button **next to** Stop whenever the draft has text) adds a message that Pi reads at its next step, which is today's steer.
  - **Drop the follow-up as a separate key.** Users who want something done afterwards say so ("after that, …"). If Thomas wants to keep it, make it a per-row toggle in the queue strip ("Send after Pi finishes") rather than a hidden chord.
  - **One wording:** "Pi will read this after its current step".
- **Why:** Analysts do not hold a mental model of steer versus follow-up, and nothing on screen teaches it. The bigger risk is the mouse: a user types a correction and clicks where Send was, but that button is now Stop. That ends the turn, possibly in the middle of a multi-step build, which is the opposite of what they meant. Comparable tools show one primary action. Cursor: Enter queues, Cmd+Enter sends now ([Cursor agent](https://cursor.com/docs/agent/overview)). Claude Code: typing and Enter is read "as soon as those [tool] calls finish, within the same turn" ([how Claude Code works](https://code.claude.com/docs/en/how-claude-code-works)), which is Pi's steer, with no second mode.
- **Cost / risk:** Low to medium. The UI is small. Removing follow-up changes `agent.followUp` wiring and the restore order. Some users may rely on Alt+Enter.

### CL-04: Make queue rows readable and individually editable

- **Today / planned:** Queue-strip rows are cut at 50 characters with no tooltip, are not focusable, and are not a live region. Restore (Alt/Option+Up) takes **everything** back and joins it into one draft, so sending it makes one message out of several ([steering and the queue › While items wait](../turn/steering-and-the-queue.md#while-items-wait), [Restoring queued messages](../turn/steering-and-the-queue.md#restoring-queued-messages)). #789 keeps this. Its "Pi didn't get to N queued messages" path merges messages the same way.
- **Proposed:** Each row shows its full text on hover and focus, and has two buttons: Edit (puts that row alone back in the composer) and Remove (×). Keep Alt/Option+Up as "edit the last row". Announce additions and deliveries politely to screen readers.
- **Why:** In a narrow pane, a 47-character prefix is often all the user can see of the instruction they are betting the next step on. "Take everything back as one blob" is the only way to fix one queued message, and it silently changes what Pi receives (two instructions become one). Cursor lets users reorder and edit queued messages ([Cursor agent](https://cursor.com/docs/agent/overview)).
- **Cost / risk:** Low. The strip and the agent queues already support removal (restore empties them). The row layout belongs to the `.pi-dock` in #793.

### CL-05: The working line says what Pi is doing, including "Stopping…"

- **Today / planned:** The working line shows one of 127 random phrases ("Herding cells into alignment…", "Consulting the void…"), changing every 6 s, plus a random hint every 4.5 s ([sending and the response › The working indicator](../turn/sending-and-the-response.md#the-working-indicator)). The composer placeholder rotates every 8 s. After Stop during a tool, Pi waits for the tool to return, with nothing changed on screen. The Stop button and the phrases stay "until the tool returns", about 11 s in the confirmed run ([stopping › At each stage](../turn/stopping.md#at-each-stage)). #789 adds a retry countdown and #793 restyles the line, but neither changes this.
- **Proposed:**
  - Show the real state: "Reading Sheet1!A1:F200", "Writing 240 cells to Summary", "Waiting for your approval", "Waiting for Chat 2 to finish a change" (the lock), "Retrying in 4s…", "Compacting context…".
  - After Stop: "Stopping after the current change…", with the Stop button disabled.
  - Keep a single neutral "Working…" while the model only streams text.
  - Stop rotating the hints and placeholders, or rotate only while idle.
- **Why:** The working line is the one place an analyst looks to judge whether to intervene, and today it carries no information. During a write it should name the target, because that is what the user may want to stop. The silent wait after Stop reads as "Stop is broken", so users press it again or close the tab. Constant motion in a narrow pane competes with the grid, and the line already truncates there (noted in #789 and #793).
- **Cost / risk:** Low to medium. Tool cards already have humanized summaries (#794 reuses them for approval wording). The playful tone is a brand choice, hence the product call.

### CL-06: A background tab that finishes, fails or needs approval is marked

- **Today / planned:** Several tabs can run at once, and the busy dot is the only mark ([sessions and tabs › Busy and locking indicators](../foundations/sessions-and-tabs.md#busy-and-locking-indicators)). "There is no sound and no system notification when a turn finishes in a background tab" ([turns › Interactions](../foundations/turns.md#interactions-with-other-systems)). #789 shows a background tab's error banner only once the user switches to it. #793 makes the dot visible and pulsing. #794 shows other tabs' approvals one at a time and names the tab.
- **Proposed:** When a background tab's turn ends, its tab shows a state until the user visits it: a filled marker for "finished" and a red marker for "failed". A tab with a pending approval gets an attention marker. The tab's accessible name includes the state.
- **Why:** Tabs exist so the user can run work in parallel. When the dot disappears, the user cannot tell finished from failed without visiting every tab. A failed overnight-style build in Chat 2 is discovered only by chance. Every chat product with parallel sessions marks unread or finished conversations.
- **Cost / risk:** Low. The state is already per tab after #789.

### CL-07: Retry a failed turn, and recall the last message

- **Today / planned:** A sent message cannot be edited, removed or recalled. "There is no history of sent drafts, no Up-arrow recall of the last message" ([the composer › After sending](../turn/the-composer.md#after-sending)). After a failure, even after #789's three automatic retries, the only way on is to type a new message. To continue after Stop, the docs suggest typing "continue" ([stopping › Stopped](../turn/stopping.md#stopped)).
- **Proposed:**
  - A failed or stopped turn shows "Retry", which resends the request from the failed point with no new user message.
  - Up in an empty composer recalls the last sent message for editing; sending it is a new message.
  - Editing history in place is not proposed (that would be a rewind). It can follow CL-01 if wanted.
- **Why:** Rate limits and network loss are routine on subscription logins (P-77). After three automatic retries the user still has to invent a prompt. Every comparable chat UI has regenerate or retry. Claude Code rewinds to a previous prompt ([checkpointing](https://code.claude.com/docs/en/checkpointing)), and ChatGPT and Cursor let users resend or edit.
- **Cost / risk:** Low for Retry, since #789 already has the retry path. Recall is a few lines in `pi-input`.

### CL-08: Tabs are titled by their conversation

- **Today / planned:** A tab shows "Chat N" unless the user renames it. The saved title (the first 80 characters of the first message) is shown only in the resume picker ([sessions and tabs › Tab labels](../foundations/sessions-and-tabs.md#tab-labels)). #795 makes numbers stable and retires closed ones, so long-lived workbooks reach "Chat 23".
- **Proposed:** After the first exchange, label the tab with a short title: the model's 3–5 word summary, or the first message trimmed. Keep a chosen title if the user renamed the tab, and use the same title in the strip, the picker, the close toast and recently closed.
- **Why:** In a 320 px pane with three to five tabs, "Chat 1 / Chat 3 / Chat 4" carries no information, so users click through tabs to find the one about the Q3 model. ChatGPT, Claude, Cursor and Copilot's history list all show conversation titles ([Copilot tips](https://support.microsoft.com/en-gb/topic/use-claude-with-agent-mode-in-excel-b2c3b3ec-154b-484b-84d0-914a80df395a)). Two names for one conversation (label and saved title) is also an inconsistency.
- **Cost / risk:** Low for a trimmed first message. A model-written title costs one small request per session, which is the product call.

### CL-09: Shift+Tab moves focus; thinking gets its own key

- **Today / planned:** Shift+Tab cycles the thinking level. #793 limits it to the composer and keeps focus there ([thinking › Choosing a level](../turn/thinking.md#choosing-a-level); #793 "Shift+Tab cycles thinking only in the composer").
- **Proposed:** In the composer, Shift+Tab moves focus backwards, as it does in every web form. Thinking stays on its status-bar control, which #793's menu model makes keyboard-reachable. If a chord is wanted, use one that browsers and Office leave free, and list it on the Shortcuts page.
- **Why:** Shift+Tab is the standard way for a keyboard or screen-reader user to leave a text field backwards. Taking it in the one field everyone uses breaks the reverse path through the taskpane (WCAG 2.1.1 expects standard keyboard operation). It also changes a costly setting by accident: an analyst tabbing back to the Files button raises reasoning one level, and every later turn is slower. The habit comes from Claude Code, where it is a terminal mode toggle with no focus order to break.
- **Cost / risk:** Trivial in code. It removes a power-user shortcut.

### CL-10: Resume always opens or focuses a tab; drop "Replace current tab" and `/resume-here`

- **Today / planned:** The picker has two target buttons, "Open in new tab" and "Replace current tab", and the hint "Default action: {action}". `/resume-here` preselects the second, and both are refused while busy ([resuming a session › Opening the picker](../sessions/resuming-a-session.md#opening-the-picker)). #795 adds "one session, one tab", so resuming an open session focuses its tab.
- **Proposed:** Remove the target buttons and `/resume-here`. Choosing a session focuses its tab if it is open; otherwise it opens in the current tab **if that tab is empty**, and in a new tab otherwise.
- **Why:** Under #795's model, "replace" only hides the current conversation from the strip, which the user then has to find again in the picker. The empty-tab case is the only one where replacing helps, and it can be automatic. One fewer command, one fewer busy refusal, and one fewer choice before every resume. Copilot's history opens a conversation with no target choice ([Copilot tips](https://support.microsoft.com/en-gb/topic/use-claude-with-agent-mode-in-excel-b2c3b3ec-154b-484b-84d0-914a80df395a)).
- **Cost / risk:** Low. Code is removed, not added.

### CL-11: Model changes while busy apply from the next turn, like thinking; drop the fork setting

- **Today / planned:**
  - Thinking can be changed while busy and applies from the next turn ("Thinking: {level} (next turn)").
  - The model cannot: `main` refuses with a toast, and #795 opens the picker read-only, with a note ([turns › Modifiers](../foundations/turns.md#modifiers)).
  - "Fork model switch into new tab" opens a copy titled "{title} ({model id})", and fork titles stack ([the model picker](../setup/the-model-picker.md)).
  - A new tab inherits the active tab's model but resets thinking to High ([sessions and tabs › What a tab owns](../foundations/sessions-and-tabs.md#what-a-tab-owns)).
- **Proposed:**
  - Accept a model choice while busy and apply it from the next turn, with the same "(next turn)" toast as thinking.
  - Remove the fork setting: switching is always in place, and Duplicate tab already makes a copy.
  - A new tab inherits the active tab's thinking level, clamped to the model (as #795 already clamps).
- **Why:** The two status-bar controls sit side by side and follow two different rules. An analyst who notices mid-turn that the tab is on the wrong model has to wait, then remember to come back. The fork setting is a global mode most users never learn, and it doubles what "choose a model" means. Resetting thinking to High in new tabs is surprising and expensive for API-key users, and it is inconsistent with inheriting the model. Copilot applies the model choice to the session ([Copilot tips](https://support.microsoft.com/en-gb/topic/use-claude-with-agent-mode-in-excel-b2c3b3ec-154b-484b-84d0-914a80df395a)).
- **Cost / risk:** Low. pi-agent-core already reads the model once per run (the same mechanism as thinking). Removing the setting needs a migration that ignores the stored key.

### CL-12: Prune and rank the command menu; refuse only what conflicts

- **Today / planned:**
  - The menu lists 26 built-ins in registration order, starting with `/model` and `/default-models`.
  - Filtering matches anywhere in the name or description, so `/res` lists `/experimental` first.
  - Arguments are dropped when a command is chosen from the menu (#789 fixes the dropping).
  - It includes `/share-session` ("Session sharing coming soon"), `/debug`, `/experimental`, and four hub aliases (`/extensions`, `/tools`, `/plugins`, `/skills`).
  - Busy rules are a hand-kept allow-list ([the command menu › Opening and filtering](../commands/the-command-menu.md#opening-and-filtering), [Running the command](../commands/the-command-menu.md#running-the-command)).
  - #795 removes `/default-models` and allows `/model`, `/settings` and `/login` while busy. #794 allows `/revert`, `/history` and `/backup`. That leaves refusals of commands that cannot conflict with a turn, such as `/name` and `/shortcuts`.
- **Proposed:**
  - Remove `/share-session` until it works.
  - Show `/debug` and `/experimental` only when typed in full.
  - Keep `/extensions` and list the three aliases under it rather than as rows.
  - Rank name-prefix matches above description matches.
  - Group the rows (Chat, Workbook, Setup).
  - Invert the busy rule: every local command runs while busy except those that would change the running tab's conversation (`/compact` stays queued).
- **Why:** The menu is how analysts discover what Pi can do. Rows that are dead ends ("coming soon") or developer tools crowd out the five commands they need (`/new`, `/resume`, `/compact`, `/history`, `/revert`) in a pane a few rows tall. The busy allow-list has already been widened twice in flight. A rule ("refuse only what conflicts") keeps it from drifting again and removes the class of bug behind P-113 and P-142. Claude Code runs status-type commands immediately and queues the rest ([commands](https://code.claude.com/docs/en/commands)).
- **Cost / risk:** Low. Mostly deletions and an ordering function. Snake (P-114) is already triaged.

## Right as planned: do not change

- **Stop returns undelivered messages to the draft and never sends them unasked** (#789). Claude Code instead sends queued messages after Esc ([interactive mode](https://code.claude.com/docs/en/interactive-mode)), but a user stops a workbook agent because something went wrong. Making them re-approve what was queued is the safe default.
- **A steer waits for the running tool batch and never interrupts a write.** Cutting a multi-range write in half would leave the workbook in a state no backup describes.
- **Stop is one action for the button, Escape and closing a tab, and queued commands are cancelled and named** (#789). Two Stops with different side effects (P-152) was the single most confusing rule in the loop.
- **Drafts are per tab and survive reload** (#789). A shared composer sent text to the wrong conversation.
- **Closing a tab is never destructive.** The session is saved, Undo is offered for 9 s, and the tab is listed in recently closed. Users close tabs freely because nothing is lost.
- **One session, one tab, with stable numbers and chosen titles** (#795). This is the basis for CL-08 and CL-10.
- **Enter never sends slash text to the model; an unknown command keeps the draft with a toast** (#789). It protects against `/` typos, and leading with a space is a clear escape hatch.
- **Transient failures retry with a visible countdown, and Stop cancels the wait** (#789). It is honest and interruptible. CL-07 only adds a manual path once the automatic retries are used up.
- **Overlays take focus, and Escape acts on the topmost one only** (#793). This removes every "one Escape did two things" defect, and it is what makes Escape-to-stop safe.
- **Changing thinking mid-turn applies from the next turn and says so.** Predictable, and the toast makes the rule visible. CL-11 extends the same rule to the model.
- **Compaction keeps the archived messages visible ("Show earlier messages") and is stoppable, per tab and saved at once** (#789). Users can still read what was summarized away.

Reviewed against `product-description` at the commit this file was added in, `main` at `ec453df`, and the open PRs #786–#795 as of 27 Oct 2026.
