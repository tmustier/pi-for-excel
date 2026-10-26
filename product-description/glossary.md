# Glossary

The vocabulary used across these documents. When a document uses one of these words, it means exactly this. Where the product has its own UI wording, the term uses it and says so.

## The taskpane

**Taskpane.** The Pi for Excel panel that Excel docks beside the grid, opened with **Open Pi** on the Home ribbon. Everything this project describes happens inside it. Excel decides its width; the user can resize it. One taskpane belongs to one workbook window. See [the taskpane](foundations/the-taskpane.md).

**Header.** The top row of the taskpane: the *tab strip* on the left and the *utilities button* (a gear) on the right. Excel draws its own title bar ("Pi for Excel", a `⋯` menu and a close button) above it; that bar is not part of the product.

**Tab strip.** The row of *session tabs* in the header, with a "+" button ("New tab") after the last tab and scroll arrows when the tabs overflow.

**Transcript.** The scrolling area under the header that shows the active tab's messages, tool cards and thinking blocks, oldest at the top. When the tab has no messages it shows the *empty state* instead.

**Empty state.** What an empty tab shows in place of the transcript: the π mark, the tagline "Understand and operate Excel. Remember your preferences. Build your own tools.", and the *starter hints*. See [empty state and hints](commands/empty-state-and-hints.md).

**Starter hint.** One of the cards in the empty state ("Explain this workbook", "Quality check this workbook", "Format this sheet", "Build my financial model"). Clicking one puts a longer prepared prompt into the composer without sending it ("Inserts into input — editable before sending.").

**Composer.** The chat input card at the bottom of the taskpane: a Files button, an auto-growing text area, and a send button that turns into a Stop button while the tab is busy. The code and the accessible name call it the chat input ("Chat message"). See [the composer](turn/the-composer.md).

**Working indicator.** The line above the composer that appears while the active tab is busy: a rotating playful phrase on the left ("Working…" first) and a rotating hint on the right ("Press Esc to stop", "Press Enter to redirect Pi", …). During a queued action such as compaction it shows a fixed label instead ("Compacting context…").

**Queue strip.** The list that appears just above the composer when the active tab has *steer*, *follow-up* or *queued* items waiting. Each row has a label ("Steer", "Follow up", "Queued", or a command such as "/compact") and the first 47 characters of the text; the last row says "↳ Press Alt+↑ to edit queued message" (⌥↑ on macOS).

**Status bar.** The row under the composer with the model control, the thinking-level control, the execution-mode control ("Auto" or "Confirm"), and the *context meter*. See [the status bar](ui/the-status-bar.md).

**Utilities menu.** The menu opened by the gear in the header ("Settings and tools"): Resume session, Rules, Extensions, Files, Backups, Settings, Keyboard shortcuts. See [the utilities menu](ui/the-utilities-menu.md).

**Tab context menu.** The menu that opens when the user right-clicks a session tab: "Rename tab…", "Duplicate tab", "Move left", "Move right", "Close other tabs", "Close tab".

**Command menu.** The pop-up list of slash commands that opens above the composer when the text starts with `/`. See [the command menu](commands/the-command-menu.md).

**Overlay.** A modal card drawn over the whole taskpane with a dimmed backdrop: settings, the model picker, the resume picker, the Backups dialog, the Files workspace, confirm dialogs, text-input dialogs. At most one overlay with a given purpose is open at a time. Escape closes most overlays. Native browser dialogs (`window.prompt`) are not overlays.

**Toast.** A short message that appears briefly at the bottom of the taskpane and disappears by itself. An **action toast** is a toast with a button, such as "Closed Chat 2 · Undo", and stays longer (9 seconds for the close toast).

**Error banner.** A red box shown above the transcript when the last run of the active tab ended in an error, such as "LLM error: …" or a context-overflow message. It is cleared when the user sends the next message, and when a later run ends without error. There is one banner for the whole taskpane, not one per tab.

**Disclosure bar.** The one-time notice above the transcript, "Pi can search the web, use extensions, and connect to external services.", with "Got it" and "Customize". It disappears once dismissed.

**Proxy banner.** The warning above the transcript, "Proxy not running · Some features won't work.", shown when the selected provider needs the CORS proxy and none is detected. See [the proxy](setup/the-proxy.md).

## Sessions and tabs

**Session.** One conversation with Pi: its messages, its model and thinking level, its title, and the workbook it is linked to. Sessions are saved in the browser storage of the add-in and can be reopened later. See [sessions and tabs](foundations/sessions-and-tabs.md).

**Session tab** (or **tab**). A session that is open in the tab strip. Each tab has its own transcript, model, thinking level, composer text, queue and busy state; switching tabs swaps all of them. Several tabs can be busy at once. A tab is not an Excel window and not a browser tab.

**Active tab.** The tab whose transcript is shown and which the composer, the status bar, Escape and slash commands act on. Exactly one tab is active.

**Default title.** The label "Chat N" that a tab shows when it has no chosen title. N is assigned when the tab is created, counting up within the taskpane's lifetime; it is not renumbered when other tabs close.

**Chosen title.** A title the user gave a tab with `/name` or the rename dialog, or that Pi gave a duplicated or forked tab ("Chat 1 copy"). A tab with a chosen title shows it instead of the default title. See [renaming a tab](sessions/renaming-a-tab.md).

**Saved title.** The title stored with a saved session and shown in the resume picker. It is the chosen title if there is one, otherwise the first 80 characters of the session's first message, otherwise "Untitled".

**Tab layout.** Which sessions are open as tabs, in what order, and which one is active. It is saved per workbook so that reopening the taskpane on the same workbook brings back the same tabs.

**Recently closed.** The list of tabs closed in this workbook that can be brought back with the undo toast, Ctrl/Cmd+Z, Ctrl/Cmd+Shift+T, `/reopen`, or the "Recently closed" section of the resume picker. See [reopening closed tabs](sessions/reopening-closed-tabs.md).

**Workbook link.** The association between a saved session and the workbook it was used in. Restore on open and the resume picker prefer sessions linked to the current workbook. The link uses a workbook identity derived by the product, never the raw file URL.

## The turn

**Turn.** Everything that happens from the moment the user sends a message until the tab is idle again: the user message joins the transcript, the model responds, tools run, the model responds again, and so on until the model stops, the user stops it, or it fails. A turn may contain many model responses and many tool calls. It is the unit of interaction in this project. See [turns](foundations/turns.md).

**Model response.** One streamed answer from the model inside a turn. It may contain text, a thinking block, and tool calls. When a model response asks for tools, the tools run and a new model response starts with their results. The code calls this a "turn"; these documents do not.

**User message.** A message the user sent, shown right-aligned in the transcript. Steer and follow-up messages become user messages when they are delivered.

**Assistant message.** A model response as it appears in the transcript, left-aligned, rendered as markdown.

**Tool call.** A request inside a model response for Pi to run one of its tools (`read_range`, `write_cells`, …) with specific inputs. Each tool call is shown as a *tool card*.

**Tool card.** The block in the transcript that shows one tool call: a header with a humanized description of what the tool is doing (for example "Read Sheet1!A1:D20"), a spinner while it runs, and a body with the input and output that can be collapsed. Consecutive related cards may be grouped. See [tool calls](turn/tool-calls.md).

**Thinking block.** The collapsible block that shows the model's reasoning, labeled "Thinking…" while it streams and "Thought for {duration}" afterwards. Only shown when the model and thinking level produce reasoning. See [thinking](turn/thinking.md).

**Idle.** A tab is idle when it is not busy: no model response is streaming, no tool is running, and no queued action is running or waiting.

**Busy.** A tab is busy from the moment a turn or a queued action (such as `/compact`) starts until the tab is idle again. A busy tab shows a small dot on its tab and, when it is the active tab, the working indicator. While busy, most slash commands are refused ("Can't run /{command} while Pi is busy"), and renaming, duplicating, changing model and closing are refused or need confirmation. What Enter and the send button do depends on whether the tab is *responding*: see below.

**Responding.** The part of busy during which a turn is running (the model is streaming or tools are running). The code calls this streaming. While responding, the send button is replaced by a Stop button, Enter steers, Alt/Option+Enter sends a follow-up, and Escape stops. A tab can be busy without responding, for example while `/compact` runs: then the send button stays, Enter queues the prompt, and neither Escape nor any button stops the compaction.

**Send.** Pressing Enter (or clicking the send button) in the composer of an idle tab with non-empty text that does not start with `/`. The text becomes a queued prompt and starts at once because the queue is empty.

**Steer.** Pressing Enter in the composer while the tab is responding. The text is delivered to Pi as a user message as soon as the current model response and its tool calls finish, before the next model response starts, and changes what Pi does next. It does not interrupt a tool that is running. The queue strip labels it "Steer". The composer placeholder while responding says "Redirect reply (↵)".

**Follow-up.** Pressing Alt/Option+Enter in the composer while the tab is responding. The text is delivered only when Pi would otherwise have finished the turn, and starts a continuation of the same turn. The queue strip labels it "Follow up". The placeholder says "New question (⌥↵)".

**Queued.** A prompt or slash command waiting in the tab's action queue to run after the current activity finishes. Prompts are queued when sent while the tab is busy but not responding (for example during `/compact`); `/compact` itself is always queued. The queue strip labels a queued prompt "Queued" and a queued command with its name.

**Restore queued.** Alt/Option+Up in the composer (or anywhere outside another text field) takes every waiting steer, follow-up and queued prompt out of the queue and puts their text back in the composer, above anything already there, separated by blank lines. Queued commands stay queued. The toast says "Restored {count} queued message(s) to editor", or "No queued messages" if there were none.

## Endings

**Stop.** The user's explicit abort of the active tab's turn: the Stop button in the composer, or Escape while the tab is responding and no menu, popover or overlay that claims Escape is open. Stop ends the current model response at once; a tool that is mid-flight is asked to cancel. What Stop does to queued items depends on how it was triggered; see [stopping](turn/stopping.md).

**Aborted.** How a turn ends after Stop. The partial model response stays in the transcript, marked "Request aborted". No error banner is shown.

**Failed.** How a turn ends when the provider or the agent loop raises an error that is not an abort. The transcript keeps everything up to the failure, and the error banner shows the reason.

**Context overflow.** A failure in which the provider rejects the request because the conversation no longer fits the model's context window. With auto-compaction on (the default), Pi compacts once and retries; see [context and compaction](turn/context-and-compaction.md).

**Interrupt.** Anything in a document's "Cancel and interrupt" table: an event that happens in the middle of an interaction and is not part of its normal path. "Interrupt" is not a product action; the product's interrupts are Stop, steer and follow-up.

## The workbook

**Workbook.** The Excel file open in the window that owns the taskpane. Pi reads and changes it only through its tools. See [the workbook](foundations/the-workbook.md).

**Mutation.** A tool call that changes the workbook: writing cells, filling formulas, inserting or deleting rows, formatting, adding comments, and so on. Tools that only read are not mutations. Some tools are a mutation or not depending on their inputs (for example `comments` reading versus adding).

**Workbook lock.** Pi's rule that only one mutation runs at a time in a taskpane, across all tabs. A tab whose mutation is waiting for another tab's shows "Locking…" on its tab. A tab whose mutation is running cannot be closed ("Wait for workbook changes to complete"). Reads are not locked.

**Execution mode.** Whether Pi applies mutations immediately (**Auto**, the default) or asks before each one (**Confirm**). The UI labels are "Auto" and "Confirm"; the command is `/yolo`, and older UI text calls Confirm "Safe mode". See [confirm mode and approvals](workbook/confirm-mode-and-approvals.md).

**Approval.** The confirm dialog Pi shows before a mutation in Confirm mode, or before a risky action in any mode. Declining makes the tool call fail with a message the model sees; the turn continues.

**Backup** (automatic). A snapshot of the cells, formats or structure a mutation is about to change, saved by Pi before the mutation runs. The UI calls these "backups" ("Snapshots saved before Pi modifies data"); the code calls them checkpoints. They are listed in the Backups dialog, can be restored one at a time, and are cleared when the workbook is saved in Excel. See [backups and revert](workbook/backups-and-revert.md).

**Revert.** Restoring the cells covered by one backup to their state before the mutation, from the Backups dialog or with `/revert` (the latest backup). Excel's own Undo does not undo Pi's changes.

**Manual backup.** A full copy of the workbook file taken on request with `/backup`, stored in the add-in's storage and downloadable later. Distinct from automatic backups. See [manual backups](workbook/manual-backups.md).

**Auto-context.** What Pi is told about the workbook before each model request without the user asking: a structural blueprint of the workbook, the current selection, and cells the user changed since the last turn. See [auto-context](turn/auto-context.md).

## Models and providers

**Provider.** A company or gateway that serves models: Anthropic, OpenAI, OpenAI Codex (ChatGPT), Google Gemini, GitHub Copilot, or a custom OpenAI-compatible gateway. A provider is **connected** once the user has logged in with OAuth or saved an API key for it. See [models and providers](foundations/models-and-providers.md).

**Model.** The specific model a tab uses (for example "GPT-6.1 Sol"). Each tab has its own model, shown in the status bar. Changing it affects only the active tab.

**Thinking level.** How much the model reasons before answering: Off, Min, Low, Medium, High, Extra high, Max, limited to the levels the model supports. Each tab has its own. Shift+Tab cycles it; a change applies from the next model request ("Thinking: {level} (next turn)").

**Context window.** The amount of conversation the model can take in at once, in tokens.

**Context meter.** The status-bar readout of how much of the context window the active tab's conversation uses, for example "1% / 500k". It warns as it fills ("Context {pct}% full.").

**Compaction.** Replacing older messages in a session with a summary written by the model, to free context. Manual with `/compact`; automatic before a prompt or between model responses when the context is nearly full (auto-compaction), and once after a context overflow. A compaction summary appears in the transcript in place of the messages it replaced.

**CORS proxy.** A small local or organisation-hosted HTTPS server that relays requests the Office webview cannot make directly, needed for some OAuth logins and providers. Configured in Settings → Proxy. See [the proxy](setup/the-proxy.md).

## Input

**Ctrl/Cmd.** The Control key on Windows and the Command key on macOS. Shortcuts in this project that say Ctrl/Cmd accept either key on either platform unless a document says otherwise.

**Alt/Option.** The Alt key on Windows and the Option key on macOS.

**Text field.** Any element the user can type into: the composer, an input in an overlay, a search box. Several shortcuts (arrow keys for tab switching, Ctrl/Cmd+Z for reopening a tab) do nothing, or behave differently, while focus is in a text field.

**Focus in the composer.** The composer's text area has keyboard focus. Escape while idle with focus in the composer moves focus out of it (to the tab strip) instead of doing anything else. F2 puts focus back in the composer from anywhere in the taskpane.
