# Glossary

The vocabulary used across these documents. When a document uses one of these words, it means exactly this. Where the product has its own UI wording, the term uses it and says so.

## The taskpane

**Taskpane.** The Pi for Excel panel that Excel docks beside the grid, opened with **Open Pi** on the Home ribbon. Everything this project describes happens inside it. Excel decides its width; the user can resize it. One taskpane belongs to one workbook window. See [the taskpane](foundations/the-taskpane.md).

**Header.** The top row of the taskpane: the *tab strip* on the left and the *utilities button* (a gear) on the right. Excel draws its own title bar ("Pi for Excel", a `⋯` menu and a close button) above it; that bar is not part of the product.

**Tab strip.** The row of *session tabs* in the header, with a "+" button ("New tab") after the last tab. Tabs retain a usable minimum width; overflow scrolls with arrows rather than shrinking tabs into slivers. The active tab is brought into view.

**Transcript.** The scrolling area under the header that shows the active tab's messages, tool cards and thinking blocks, oldest at the top. When the tab has no messages it shows the *empty state* instead.

**Empty state.** What an empty tab shows in place of the transcript: the π mark, the tagline "Understand and operate Excel. Remember your preferences. Build your own tools.", and the *starter hints*. See [empty state and hints](commands/empty-state-and-hints.md).

**Starter hint.** One of the cards in the empty state ("Explain this workbook", "Quality check this workbook", "Format this sheet", "Build my financial model"). Clicking one puts a longer prepared prompt into the composer without sending it ("Inserts into input — editable before sending.").

**Composer.** The chat input card at the bottom of the taskpane: a Files button, an auto-growing text area, and a send button that turns into a Stop button while the tab is busy. The code and the accessible name call it the chat input ("Chat message"). See [the composer](turn/the-composer.md).

**Working indicator.** The line above the composer that appears while the active tab is busy: a rotating playful phrase on the left ("Working…" first) and a rotating hint on the right ("Press Esc to stop", "Press Enter to redirect Pi", …). During a queued action such as `/compact` or auto-compaction before a prompt it shows a fixed label instead ("Compacting context…"); auto-compaction between model responses keeps the rotating phrases.

**Queue strip.** The list just above the composer of waiting steers, follow-ups, queued actions and held "Not sent" messages. Each message has its own Edit and ×; full text is available on hover or focus. Alt/Option+Up edits the newest message row, not every row at once. Not sent rows survive reload and never send themselves.

**Status bar.** The row under the composer with, in order, the model control, the thinking-level control, the *context meter*, the lock badge (while a mutation holds the workbook lock) and the execution-mode control ("Auto" or "Confirm"). See [the status bar](ui/the-status-bar.md).

**Status popover.** The small dialog that opens above a status-bar control when clicked: the thinking popover ("Thinking level") or the context popover ("Context usage", with "/compact" and "/new" buttons). Only one is open at a time. It closes on Escape, an outside click, clicking its control again, or when the model picker opens. Escape goes to an open popover first, so it never stops a turn.

**Lock badge.** The status-bar badge shown while the active tab waits for ("lock…") or holds ("lock") the workbook lock; the status-bar counterpart of the tab's "Locking…" label. English only.

**Utilities menu.** The menu opened by the gear in the header ("Settings and tools"), in order: Settings, Extensions, Files, Rules, Resume session, Backups, Keyboard shortcuts. See [the utilities menu](ui/the-utilities-menu.md).

**Tab context menu.** The menu that opens when the user right-clicks a session tab: "Rename tab…", "Duplicate tab", "Move left", "Move right", "Close other tabs", "Close tab".

**Command menu.** The pop-up list of slash commands that opens above the composer when the text starts with `/`. See [the command menu](commands/the-command-menu.md).

**Overlay.** A modal card drawn over the whole taskpane with a dimmed backdrop: settings, the model picker, the resume picker, the Backups dialog, the Files workspace, confirm dialogs, text-input dialogs. At most one overlay with a given purpose is open at a time. Escape closes every overlay built on the shared overlay dialog (resume picker, Backups, confirm and text-input dialogs, settings, the model picker), and while one is open Escape does not stop a turn. Native browser dialogs (`window.prompt`) are not overlays.

**Toast.** A temporary message above the taskpane, including overlays. Toasts stack, duplicates coalesce, and hover pauses their expiry. An **action toast** offers a button, such as "Closed Chat 2 · Undo". Errors are dismissible and announced assertively; ordinary messages are polite.

**Error banner.** A dismissible red notice below the tabs for the active tab's failure. Each tab owns its banner; switching shows that tab's error or none, including failures that happened in the background. A new send clears the tab's error.

**Disclosure bar.** The one-time notice above the transcript, "Pi can search the web, use extensions, and connect to external services.", with "Got it" and "Review in Settings". It disappears once dismissed.

**Proxy banner.** A dismissible notice shown when the enabled proxy cannot be reached. It is absent with the proxy disabled. On Excel on the web it explains when the browser prevents add-ins reaching the local proxy. See [the proxy](setup/the-proxy.md).

**Loading screen.** The startup spinner. A slow start shows "Still starting…" after 12 seconds and stronger guidance at 60 seconds, but can still mount normally. Only an actual failure becomes a startup error.

**Startup error.** A fatal initialization failure, distinct from a slow-start notice and a tab's turn-error banner. Close and reopen Pi after resolving its cause.

**Extension widget slot.** The space between the working indicator and the queue strip where an installed extension can show a widget. Hidden in the default configuration. While it holds a widget it claims Escape before Stop.

**Settings overlay.** The shared Settings shell and page stack. Escape goes back along the path taken, and closes at the page where the shell opened. The up button goes to a parent page. All modals take and trap focus and restore it when closed.

## Sessions and tabs

**Session.** One conversation with Pi: its messages, its model and thinking level, its title, and the workbook it is linked to. Sessions are saved in the browser storage of the add-in and can be reopened later. See [sessions and tabs](foundations/sessions-and-tabs.md).

**Session tab** (or **tab**). A session open in the tab strip. Each tab owns its transcript, model, thinking level, queue, draft, Not sent rows, error and busy state; switching swaps them all. Several tabs can be busy at once, but a session is open in only one tab in a taskpane. A tab is not an Excel window or browser tab.

**Active tab.** The tab whose transcript is shown and which the composer, the status bar, Escape and slash commands act on. Exactly one tab is active.

**Default title.** The label "Chat N" that a tab shows when it has no chosen title. N is assigned when the tab is created, counting up within the taskpane's lifetime; it is not renumbered when other tabs close. Duplicating a tab consumes a number.

**Chosen title.** A title the user gave a tab with `/name` or the rename dialog, or that Pi gave a duplicated or forked tab ("Chat 1 copy"). A tab with a chosen title shows it instead of the default title. See [renaming a tab](sessions/renaming-a-tab.md).

**Saved title.** The title stored with a saved session and shown in the resume picker. It is the chosen title if there is one, otherwise the first 80 characters of the session's first message, otherwise "Untitled".

**Tab layout.** Which sessions are open as tabs, in what order, and which one is active. It is saved per workbook so that reopening the taskpane on the same workbook brings back the same tabs.

**Recently closed.** The list of tabs closed in this workbook that can be brought back with the undo toast, Ctrl/Cmd+Z, Ctrl/Cmd+Shift+T, `/reopen`, or the "Recently closed" section of the resume picker. See [reopening closed tabs](sessions/reopening-closed-tabs.md).

**Workbook link.** The association between a saved session and the workbook it was used in. Restore on open and the resume picker prefer sessions linked to the current workbook. The link uses a workbook identity derived by the product, never the raw file URL.

**Workbook identity.** How Pi recognises a workbook across reopenings: a SHA-256 hash of the address Excel reports for the file (its OneDrive or SharePoint URL, or its path on disk) with any query string or fragment dropped. The raw address is never stored. The identity is unknown when there is no address: a never-saved workbook on desktop, the standalone taskpane, a host that does not answer within 3 s, and Excel on the web when the workbook is opened signed out through a shared link (confirmed). With no identity the resume picker has no workbook filter, sessions stay unlinked, and all such workbooks share one tab layout. Renaming, moving or Save As gives a new identity. Pi re-reads it whenever the taskpane gains focus or becomes visible. Tab layouts, workbook links and the resume picker's filter all key on it.

**Latest session.** The session most recently saved or loaded for a workbook. When no tab layout exists, opening the taskpane restores it; if the workbook identity is unknown, Pi restores the most recently modified session from any workbook instead (the global fallback).

**Resume picker.** The "Resume session" overlay opened by `/resume`, `/resume-here` or "Resume session" in the utilities menu: two target buttons ("Open in new tab", "Replace current tab"), a "Show sessions from all workbooks" filter, a "Recently closed" section (up to 6 entries) and "Saved sessions" (up to 30). See [resuming a session](sessions/resuming-a-session.md).

**Busy dot.** The small dot after a tab's title while that tab is busy. It is hidden from screen readers.

## The turn

**Turn.** Everything that happens from the moment the user sends a message until the tab is idle again: the user message joins the transcript, the model responds, tools run, the model responds again, and so on until the model stops, the user stops it, or it fails. A turn may contain many model responses and many tool calls. It is the unit of interaction in this project. See [turns](foundations/turns.md).

**Model response.** One streamed answer from the model inside a turn. It may contain text, a thinking block, and tool calls. When a model response asks for tools, the tools run and a new model response starts with their results. The code calls this a "turn"; these documents do not.

**User message.** A message the user sent, shown right-aligned in the transcript. Steer and follow-up messages become user messages when they are delivered.

**Assistant message.** A model response as it appears in the transcript, left-aligned, rendered as markdown.

**Tool call.** A request inside a model response for Pi to run one of its tools (`read_range`, `write_cells`, …) with specific inputs. Each tool call is shown as a *tool card*.

**Tool card.** The block in the transcript that shows one tool call: a header with a humanized description of what the tool is doing (for example "Read Sheet1!A1:D20"), a spinner while it runs, and a body with the input and output that can be collapsed. Consecutive related cards may be grouped. See [tool calls](turn/tool-calls.md).

**Tool group.** A container for related consecutive tool cards in one model response, headed with a live count ("3 reads", "4 edits"). Two cards start expanded; three or more start collapsed unless the user chose otherwise. Collapsing hides the cards. Each card finishes independently.

**Details mode.** The remembered taskpane-wide state Ctrl/Cmd+O toggles ("Details expanded" / "Details collapsed", with the platform shortcut). It starts collapsed and applies to cards, groups and thinking blocks, including new content, rerenders and other tabs. A manual toggle overrides one item until the next global toggle.

**Thinking block.** Collapsible reasoning, labelled "Thinking…" only while reasoning streams, then "Thought for {duration}" when its start was observed, otherwise "Thought". Its header is a keyboard-operable button. See [thinking](turn/thinking.md).

**Idle.** A tab is idle when it is not busy: no model response is streaming, no tool is running, and no queued action is running or waiting.

**Busy.** A tab is busy from the moment a turn or a queued action (such as `/compact`) starts until the tab is idle again. A busy tab shows a small dot on its tab and, when it is the active tab, the working indicator. While busy, most slash commands are refused ("Can't run /{command} while Pi is busy"), and renaming, duplicating, changing model and closing are refused or need confirmation. What Enter and the send button do depends on whether the tab is *responding*: see below.

**Responding.** The part of busy in which a turn runs (streaming or tools). Enter steers and Alt/Option+Enter follows up. A tab can be busy without responding, such as during compaction: Enter queues a prompt. Stop is available for both turn and compaction work.

**Send.** Pressing Enter (or clicking the send button) in the composer of an idle tab with non-empty text that does not start with `/`. The text becomes a queued prompt and starts at once because the queue is empty.

**Steer.** Pressing Enter in the composer while the tab is responding. The text is delivered to Pi as a user message as soon as the current model response and its tool calls finish, before the next model response starts, and changes what Pi does next. It does not interrupt a tool that is running. The queue strip labels it "Steer". The composer placeholder while responding says "Redirect reply (↵)".

**Follow-up.** Pressing Alt/Option+Enter in the composer while the tab is responding. The text is delivered only when Pi would otherwise have finished the turn, and starts a continuation of the same turn. The queue strip labels it "Follow up". The placeholder says "New question (⌥↵)".

**Queued.** A prompt or slash command waiting in the tab's action queue to run after the current activity finishes. Prompts are queued when sent while the tab is busy but not responding (for example during `/compact`); `/compact` itself is always queued. The queue strip labels a queued prompt "Queued" and a queued command with its name.

**Restore queued.** Alt/Option+Up edits the newest editable queue row, putting only its text above the existing draft. Other messages and commands stay. Edit on an individual row targets that row. With none, the toast says "No queued messages to restore".

**Busy-allowed command.** A command accepted while a tab works. This includes model/settings/login, copy/export and backup/revert as well as the queue, navigation and configuration commands listed in [the command menu](commands/the-command-menu.md). A busy model picker is read-only. Refused commands keep their draft.

## Endings

**Stop.** The user's explicit abort of the active tab's turn: the Stop button in the composer, or Escape while the tab is responding and no menu, popover or overlay that claims Escape is open. Stop ends the current model response at once; a tool that is mid-flight is asked to cancel. What Stop does to queued items depends on how it was triggered; see [stopping](turn/stopping.md).

**Aborted.** How a turn ends after Stop. The partial model response stays in the transcript, marked "Request aborted". No error banner is shown.

**Failed.** How a turn ends when the provider or the agent loop raises an error that is not an abort. The transcript keeps everything up to the failure, and the error banner shows the reason.

**Context overflow.** A failure in which the provider rejects the request because the conversation no longer fits the model's context window. With auto-compaction on (the default), Pi compacts once and retries; see [context and compaction](turn/context-and-compaction.md).

**Stranded.** The historical defect where undelivered text leaked into the next turn. Current behavior holds that text in separate saved Not sent rows and clears it from the delivery queues.

**Interrupt.** Anything in a document's "Cancel and interrupt" table: an event that happens in the middle of an interaction and is not part of its normal path. "Interrupt" is not a product action; the product's interrupts are Stop, steer and follow-up.

## The workbook

**Workbook.** The Excel file open in the window that owns the taskpane. Pi reads and changes it only through its tools. See [the workbook](foundations/the-workbook.md).

**Mutation.** A tool call that changes the workbook: writing cells, filling formulas, inserting or deleting rows, formatting, adding comments, and so on. Tools that only read are not mutations. Some tools are a mutation or not depending on their inputs (for example `comments` reading versus adding).

**Workbook lock.** Serializes mutations across tabs in one taskpane. Waiting tabs show "Locking…". Reads in the same runtime wait for earlier registered mutations; reads in another runtime do not. Stop and close waits for submitted workbook work to settle.

**Execution mode.** Whether Pi applies mutations immediately (**Auto**, the default) or asks before each one (**Confirm**). The UI labels are "Auto" and "Confirm"; the command is `/yolo`, and older UI text calls Confirm "Safe mode". See [confirm mode and approvals](workbook/confirm-mode-and-approvals.md).

**Approval.** The confirm dialog Pi shows before a mutation in Confirm mode, or before a risky action in any mode. Declining makes the tool call fail with a message the model sees; the turn continues.

**Backup** (automatic). A saved pre-mutation snapshot listed in Backups and restorable independently of Excel Undo. Saving and reloading do not clear it; retention and explicit deletion do. See [backups and revert](workbook/backups-and-revert.md).

**Revert.** Restoring the cells covered by one backup to their state before the mutation, from the Backups dialog or with `/revert` (the latest backup). On Excel on the web, Excel's own Undo also takes back Pi's changes step by step, and Ctrl+Y redoes them (confirmed on the web and on Excel for Mac). A revert restores a specific backup regardless of Undo history.

**Manual backup.** A full copy of the workbook file taken on request with `/backup`, stored in the add-in's storage and downloadable later. Distinct from automatic backups. See [manual backups](workbook/manual-backups.md).

**Auto-context.** What Pi is told about the workbook before each model request without the user asking: a structural blueprint of the workbook, the current selection, and cells the user changed since the last turn. See [auto-context](turn/auto-context.md).

**Audit log.** Pi's local record of every workbook mutation it made or tried, and of the user's restores: tool, workbook, range, changed cells with before and after values, whether it was blocked, execution mode. It keeps the newest 500 entries across all workbooks, is not shown in the UI (except through the "Explain these changes" drawer on mutation tool cards), is not cleared on save, and is exported with `/export audit`. It is not a backup and cannot revert anything.

**Cell link.** A blue cell or range address inside a tool card (title, humanized input, change tables, dependency trees). Clicking it makes Excel activate the sheet and select the range (the first area of a multi-area address) without changing the workbook. Addresses in Pi's own prose are not cell links.

**Rules.** Instructions for Pi, written by the user on the Rules page or by Pi through its `instructions` tool. User rules apply everywhere; workbook rules apply to one identified workbook. They take effect from the next request.

**Conventions.** Pi's default number-format preferences (currency, negatives, zeros, decimal places), set with the `conventions` tool or on the Rules page.

**Skill.** A task playbook (`SKILL.md`) Pi loads on demand with its `skills` tool. Bundled skills ship with the add-in; external skills are installed by the user into the Files workspace under `skills/`.

**Files workspace.** Pi's own file store in add-in storage, or a connected folder on disk. Shared by every tab and workbook on this computer, not part of the workbook, and read by Pi with its `files` tool, never as an attachment. Includes read-only built-in docs under `assistant-docs/`. Files imported are tagged to the workbook that was open (the "Workbook" badge), but only when Pi can identify the workbook; on an unidentified workbook they are untagged.

**Extension.** Add-on code (also called a plugin) that registers tools, slash commands, widgets or connections. Code and URL installs run in an isolated sandbox frame; built-in and local ones run in the taskpane. Managed on the "Plugins" page; the bundled Snake extension is installed and enabled by default.

**External tools.** Pi's network tools: `web_search` with `fetch_page`, and `mcp`. One "External tools" switch on the Connections page turns them on or off. They never ask for approval.

## Models and providers

**Provider.** A company or gateway that serves models: Anthropic, OpenAI, OpenAI Codex (ChatGPT), Google Gemini, GitHub Copilot, or a custom OpenAI-compatible gateway. A provider is **connected** once the user has logged in with OAuth or saved an API key for it. See [models and providers](foundations/models-and-providers.md).

**Model.** The specific model a tab uses (for example "GPT-6.1 Sol"). Each tab has its own model, shown in the status bar. Changing it affects only the active tab.

**Thinking level.** How much the model reasons before answering: Off, Min, Low, Medium, High, Extra high, Max, limited to the levels the model supports. Each tab has its own. Shift+Tab cycles it; a change applies from the next turn, never within the running one ("Thinking: {level} (next turn)").

**Context window.** The amount of conversation the model can take in at once, in tokens.

**Context meter.** The status-bar readout of how much of the context window the active tab's conversation uses, for example "1% / 500k". It warns as it fills ("Context {pct}% full.").

**Compaction.** Replacing older messages in a session with a summary written by the model, to free context. Manual with `/compact`; automatic before a prompt or between model responses when the context is nearly full (auto-compaction), and once after a context overflow. A compaction summary appears in the transcript in place of the messages it replaced.

**Compaction summary card.** The collapsed "Summarized {n} tokens" card that stands in for compacted messages; it expands to "Summary".

**Archived history.** The collapsed "Show earlier messages" card at the top of a compacted session holding every message compacted away. Saved with the session, never sent to the model.

**Auto-compaction thresholds.** For windows of 200k tokens or more, 85% of the window; 128k–199k, the smaller of 88% and the window minus 16,384; under 128k, the window minus the smaller of 16,384 and half the window. The tab needs at least 4 messages. Compaction keeps the most recent ~20,000 tokens unchanged.

**CORS proxy.** A small local or organisation-hosted HTTPS server that relays requests the Office webview cannot make directly, needed for some OAuth logins and providers. Configured in Settings → Proxy. See [the proxy](setup/the-proxy.md).

**Credential.** What makes a provider connected: an OAuth login (access and refresh token), a saved API key, or a saved custom gateway (even one with no key). Stored per machine in the add-in's browser storage, never in the workbook. Saving a key does not test it. OAuth logins are restored, and refreshed if expired, only when the taskpane opens. Disconnecting deletes the stored credential without revoking it at the provider.

**Default model.** The model Pi chooses for the first tab of a fresh taskpane and as the replacement when a tab's provider stops being connected; recomputed when the set of connected providers changes. Preference: OpenAI (API), then OpenAI (ChatGPT), then Anthropic's newest Opus, then Google, then any provider's newest; with nothing connected, GPT-6.1 Sol from OpenAI (API). New tabs copy the active tab's model, not the default.

**Connection dialog.** The overlay "Connect {label}" ("This model needs a login or API key to continue.") that a turn opens when its model's provider has no credential. Connecting continues the turn; dismissing fails it.

**Welcome overlay.** The first-run overlay ("Pi for Excel", "Connect an AI provider to get started", "Select provider") shown when the taskpane opens with no provider connected. It closes on any connection; Escape or a backdrop click dismiss it. It does not take focus, so with focus in the composer the first Escape only leaves the composer.

**Model picker.** The "Select model" overlay, opened by the status-bar model control, `/model` or `/default-models`.

**Model-switch behavior.** What choosing a model does to a tab with messages: **in place** (default) changes the tab's model; **fork** (Settings: "Fork model switch into new tab") opens a copy titled "{source title} ({model id})".

**Proxy gate.** The dialog "One more step before login" shown before an OAuth login that needs the CORS proxy when none answers, with "Copy", "Retry" and "Cancel".

**Custom gateway.** A user-added OpenAI-compatible endpoint. It becomes a provider named "Gateway · {name}" with exactly one model.

**Add-in storage.** The browser storage of the add-in's web origin on this computer: the `pi-for-excel` IndexedDB database (sessions, settings, keys, backups), a few `localStorage` flags, and the private file area of the Files workspace. Every taskpane on the same host and browser profile shares it; Excel desktop and Excel on the web never share it. Clearing site data erases all of it. See [persistence and restore](cross-cutting/persistence-and-restore.md).

**Save boundary.** Excel reporting a save. It does not clear Pi's automatic backups, on desktop or the web.

**Document token.** A random identifier Pi writes into a never-saved workbook (add-in setting `pi.workbookInstanceId`) before its first automatic backup, used to file backups until the workbook has a workbook identity.

## Input

**Ctrl/Cmd.** The Control key on Windows and the Command key on macOS. Shortcuts in this project that say Ctrl/Cmd accept either key on either platform unless a document says otherwise.

**Alt/Option.** The Alt key on Windows and the Option key on macOS.

**Text field.** Any element the user can type into: the composer, an input in an overlay, a search box. Several shortcuts (arrow keys for tab switching, Ctrl/Cmd+Z for reopening a tab) do nothing, or behave differently, while focus is in a text field.

**Draft.** Unsent composer text belonging to one tab and saved with its session. Switching swaps drafts; hints append without sending and can be undone. Stop holds undelivered messages separately rather than joining them to the draft.

**Auto-scroll.** The transcript following new content to the bottom while Pi responds. Scrolling up more than about 32 px from the bottom turns it off; scrolling back within about 20 px, or sending a message, turns it on again.

**Focus in the composer.** The composer's text area has keyboard focus. Escape while idle with focus in the composer moves focus out of it (to the tab strip) instead of doing anything else. F2 puts focus back in the composer from anywhere in the taskpane.
