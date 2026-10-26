# Pi for Excel product description

A written description of the user experience of Pi for Excel: what the user sees, what they can do, and exactly what happens when they do it.

## Purpose

Pi for Excel is, from the user's point of view, a large state chart. The user moves through it with typed messages, Enter and its modifiers, Escape, clicks on tabs, buttons, status-bar controls and menus, slash commands, dropped files, and the edits and selection changes they make in the workbook itself. Most of that behavior is defined implicitly, spread across the taskpane bootstrap (`src/taskpane/`), the agent runtime from pi-agent-core, the sidebar components (`src/ui/`), the slash commands (`src/commands/`), the workbook tools (`src/tools/`), and the tests. There is no single place that says, in plain language, "when the user does X, this is what happens, and this is what happens if they do Y halfway through."

This project is that place. It describes the full experience a user has in the Pi taskpane inside Excel (desktop or Excel on the web), with one model provider connected, in the default configuration: no extensions installed, no experimental features enabled, no custom rules or conventions, execution mode at its default.

The documents are for people who need to understand or change the product: designers, engineers, writers, testers, and anyone evaluating whether a behavior is intentional. They are written from the outside in. They describe the experience, not the implementation.

### What this is not

- Not API documentation. Extension authoring lives in [docs/extensions.md](../docs/extensions.md); bridge contracts live in [docs/tmux-bridge-contract.md](../docs/tmux-bridge-contract.md) and [docs/python-bridge-contract.md](../docs/python-bridge-contract.md); tool behavior rules for the model live in [src/tools/DECISIONS.md](../src/tools/DECISIONS.md).
- Not organized by module. `src/taskpane`, `src/ui`, `src/commands` and `src/tools` are not described separately. A single behavior is described once, wherever the user encounters it.
- Not a technical design document. Where a technical detail is critical to understanding the experience, it appears in a block quote labeled `Technical note:` and nowhere else.
- Not a description of what the model says. Responses are not deterministic; documents describe the structure, timing and state of a turn, never its wording.

## Conventions

- Describe the experience, not the code. "The message waits in the queue and is sent when Pi finishes" rather than "the action queue awaits `agent.waitForIdle()`".
- Technical detail goes in block quotes, prefixed with `Technical note:`. Use it only when the mechanism changes what the user would expect.
- Use sentence case for headings.
- Name the vocabulary consistently. The [glossary](glossary.md) is the source of truth for terms like *turn*, *session*, *tab*, *busy*, *steer*, *follow-up*, *queued*, *execution mode* and *backup*.
- Use the product's own UI wording for controls and messages, quoted exactly ("Wait for this tab to finish before renaming").
- Keys are written as Windows/macOS pairs where they differ: Ctrl/Cmd, Alt/Option. Enter is written Enter.
- Every document ends with the commit of pi-for-excel it was verified against and a list of open questions.
- When a behavior is surprising, say so and say why it is that way if the reason is known. Do not smooth it over.

## The work to be done

Each document describes one feature. Features are large things (sending a message and watching the response stream, with tools running in the middle) or small things (renaming a tab), but each is described in full, including its edge cases and its interactions with other features.

### Document template

Every feature document follows the same skeleton so that documents are comparable and nothing is skipped. The full skeleton, with guidance for each section, is in [goal.md](goal.md#document-skeleton).

1. **Summary.** One paragraph describing the feature abstractly. For example: "Renaming gives a session tab a title the user chose instead of the one Pi generated."
2. **The simple case.** The common path in prose.
3. **The interaction, event by event.** The five phases of a *turn* (or of the feature's own interaction, mapped onto the same five slots): **compose** (what the input or dialog accepts and what is decided before anything is sent or saved); **ends at once** (empty input, a local command, a cancelled dialog: nothing reaches the model or the workbook); **sent** (the moment the action is committed and aborting is no longer free); **while busy** (what updates live, what the user can still do); **finishes** (what is persisted, where the user lands, the failure path). Include a small state diagram (Mermaid `stateDiagram-v2`) of the states the user passes through.
4. **Modifiers.** A table of the product's variant axis, with the same rows in every document, and what each does when set before the interaction and when changed while busy:
    - The model and provider selected for the tab
    - The thinking level
    - The execution mode (Auto or Confirm)
    - Whether files are in the Files workspace, or a dialog or overlay is open
    - The tab's state: new and empty, restored with history, or another tab busy
    - The host: Excel desktop or Excel on the web
5. **Cancel and interrupt.** The same checklist in every document, in this order:
    - Stop: the Stop button or Escape
    - The user sends something else: Enter (steer), Alt/Option+Enter (follow-up), a slash command, or a starter hint
    - The user switches, creates, closes or reopens a tab
    - An overlay or menu opens over the interaction: settings, the model picker, resume, a confirm dialog, the command menu
    - The model or thinking level changes
    - The provider fails: network lost, request errors, rate limit, login expired
    - The taskpane goes away: reloaded, closed, the workbook or Excel closed
    - The workbook changes underneath: the user edits cells or moves the selection, another tab holds the workbook lock, the workbook is saved
    - Focus moves to the grid or another window
    - The context window fills: auto-compaction or an overflow
6. **Interactions with other systems.** In this fixed order: **Persistence.** **Workbook changes and recovery.** **Other tabs and sessions.** **Execution mode and approvals.** **Models and providers.** **Offline and provider errors.** **Notifications.** **Settings and preferences.** **Host differences.** **Language.** **Accessibility and keyboard.**
7. **Edge cases.** Anything a user could notice that is not covered above.
8. **Open questions and verification.** The pi-for-excel commit the document was verified against, and any behavior that could not be confirmed.

Item 5 matters most. Asking the same interrupt questions of every feature is how gaps and inconsistencies are found.

### Method

For each document:

1. Read the interaction state for the feature: `src/taskpane/` (bootstrap, keyboard shortcuts, action queue, runtime manager, status bar) and the owning component in `src/ui/` or command in `src/commands/builtins/`.
2. Read the matching tests in `tests/`. Files like `action-queue.test.ts`, `session-runtime-manager.test.ts`, `tab-layout.test.ts`, `session-title.test.ts`, `auto-compaction.test.ts`, `overflow-recovery.test.ts`, `workbook-write-inspect-undo-contract.test.ts` and `status-popovers.test.ts` are close to executable specifications of the edge cases.
3. Draft the document.
4. Try anything ambiguous in the running product: Pi in Excel on the web from an orb (`node scripts/orb/excel-web-session.mjs --check`, see [docs/orb-excel-web.md](../docs/orb-excel-web.md)), or the taskpane alone at `https://localhost:3141/src/taskpane.html`, which boots without Excel after a 3-second Office timeout. Tests settle "what happens"; the running product settles how it feels, what is visible while the interaction is in progress, and what the timing is like.
5. Record the commit verified against.

### Verification

Drafting reads the code; verification watches the product. The `verification/` directory holds one checklist per cluster of documents, each item a single observable claim with setup, steps, expected result, a priority, and the device it needs. A tester runs them in Excel, records `pass`, `fail`, or `blocked` in the Result column, and files every failure in `bug-triage.md` with the item's ID. A document moves from `drafted` to `verified` in the coverage table only when every P1 and P2 item for it has passed or been filed.

`bug-triage.md` is the other half: every behavior the documents flagged as a likely defect, deduplicated, with reproduction steps, the reason in the code, a severity, and the decision the product team needs to make. Entries confirmed in the running product carry a Status line.

### Order of work

1. **Pilot: renaming a tab.** Small and self-contained. Used to settle the template, tone, and depth.
2. **Foundations: the taskpane, turns, sessions and tabs, the workbook, models and providers.** Everything else refers to them.
3. **The turn.** The composer, the streaming response, tool calls, steering and the queue, Stop, thinking, errors. The bulk of the experience and the hardest part. Written third so the template is already proven.
4. **Everything else.** Once the template and two exemplars exist, the remaining documents can be drafted in parallel, followed by a consistency pass and a verification pass across the whole set.

Progress is tracked in the coverage table below.

### Scope decisions

- **WPS Spreadsheets.** Excluded. WPS is a second host with its own JSAPI and its own tool set; see [docs/wps-support.md](../docs/wps-support.md). It can get its own description later.
- **Extension authoring and the extension runtime.** Excluded beyond what a user meets in the Extensions hub (installing, enabling, removing). What an installed extension can do is defined by the extension.
- **Experimental features and local bridges.** `/experimental`, the tmux bridge, the Python and LibreOffice bridge, direct Office.js execution and the debug UI (`/debug`) are off by default. They get one document, [extensions/experimental.md](extensions/experimental.md), which says what turning each on changes, and are otherwise out of scope.
- **Installing the add-in.** Sideloading and manifests are covered by [docs/install.md](../docs/install.md). The description starts when the taskpane opens.
- **Keyboard shortcuts are described inside each document** that owns the behavior, and listed once in [ui/keyboard-shortcuts.md](ui/keyboard-shortcuts.md) as an index that links back. A separate behavioral document would drift.
- **Persistence is described once** in [cross-cutting/persistence-and-restore.md](cross-cutting/persistence-and-restore.md) and linked from the Persistence paragraph of every document, because it is one mechanism (browser storage scoped to the add-in origin) with one set of failure modes.
- **Interaction shape.** The unit of interaction is a turn and its phases are compose, ends at once, sent, while busy, finishes. The variant rows, the interrupt list and the order of cross-cutting concerns are fixed as written in the document template above.
- **Numbered rules.** These are prose documents, not numbered specifications. Stable heading anchors are enough for cross-references.
- **Location.** This project lives in `product-description/` at the repository root, not under `docs/`. Every markdown file under `docs/` (outside the archive and smoke-run logs) is bundled into the taskpane as read-only "assistant-docs" (`src/files/builtin-docs.ts`), and these documents are not meant to ship in the add-in.

## Structure

```
README.md                        this file
goal.md                          the standing instructions for whoever drafts
AGENTS.md, CLAUDE.md             entry points for agents: read README.md, then goal.md
glossary.md                      shared vocabulary
bug-triage.md                    suspected defects collected from every document, with repro steps and decisions needed

verification/
  README.md                      how to run a hand-verification pass and record results
  foundations-and-turn.md        checklists for foundations/ and turn/
  sessions-and-commands.md       checklists for sessions/ and commands/
  workbook.md                    checklists for workbook/
  setup-guidance-files.md        checklists for setup/, guidance/ and files/
  extensions-ui-cross-cutting.md checklists for extensions/, ui/ and cross-cutting/

foundations/
  the-taskpane.md                the regions of the sidebar, what is shown on arrival, and how
                                 overlays, toasts and banners layer over it
  turns.md                       what a turn, message, tool call and busy are; steer, follow-up
                                 and queue; what Stop does at each stage
  sessions-and-tabs.md           what a session and a tab are, how they relate to the workbook,
                                 and what is restored when the taskpane reopens
  the-workbook.md                what Pi can read and change, the workbook lock, automatic backups,
                                 and the execution mode
  models-and-providers.md        providers, logins and API keys, the model of a tab, and
                                 thinking levels

turn/
  the-composer.md                the chat input: typing, Enter and Shift+Enter, placeholders,
                                 the Files button, dropping files
  sending-and-the-response.md    the simple turn from Enter to the finished response, including
                                 the working indicator and the error banner
  tool-calls.md                  tool cards while a tool runs and after, grouping, and showing or
                                 hiding details with Ctrl/Cmd+O
  steering-and-the-queue.md      sending while Pi is busy: steer, follow-up, queued prompts and
                                 commands, and restoring them with Alt/Option+Up
  stopping.md                    Stop and Escape at every stage of a turn
  thinking.md                    the thinking level, Shift+Tab, and thinking blocks
  context-and-compaction.md      the context meter, /compact, auto-compaction and overflow recovery
  auto-context.md                what Pi is told about the workbook, the selection and recent edits
                                 before each turn

sessions/
  renaming-a-tab.md              the pilot: double-click, "Rename tab…" and /name
  tabs.md                        creating, switching, moving, duplicating and closing tabs
  reopening-closed-tabs.md       Ctrl/Cmd+Z, Ctrl/Cmd+Shift+T, /reopen and the undo toast
  resuming-a-session.md          /resume, /resume-here and the resume picker
  export-and-copy.md             /export, /copy and /share-session

commands/
  the-command-menu.md            typing /, filtering, choosing, and the busy policy for commands
  empty-state-and-hints.md       the empty tab, its tagline and the starter hints

workbook/
  reading.md                     overview, read, search, trace and explain: what the user sees
                                 when Pi reads
  writing.md                     writing cells, filling formulas, structure, formatting,
                                 conditional formats, view settings, comments, charts and names
  confirm-mode-and-approvals.md  Auto and Confirm, /yolo, and approval dialogs
  backups-and-revert.md          automatic backups, the Backups dialog, /history and /revert
  manual-backups.md              /backup: create, list, restore and clear full-workbook copies

setup/
  welcome-and-login.md           the first-run screen, connecting a provider, OAuth and API keys
  the-model-picker.md            /model, the status-bar model control, Ctrl+P and /default-models
  settings.md                    /settings and its pages
  custom-gateways.md             adding an OpenAI-compatible gateway
  the-proxy.md                   the CORS proxy setting, the proxy banner and its polling

guidance/
  rules.md                       /rules: user-level and workbook-level instructions
  conventions.md                 formatting conventions Pi follows
  skills.md                      /skills and the bundled agent skills

files/
  the-files-workspace.md         /files, importing, browsing and what Pi can read or write

extensions/
  the-extensions-hub.md          /extensions: connections, plugins and skills tabs
  web-search-and-mcp.md          /tools: web search, fetch page and MCP servers
  experimental.md                /experimental and what each flag changes

ui/
  the-status-bar.md              the model, thinking, mode and context controls and their popovers
  message-rendering.md           markdown, code blocks, tables, cell links and attachments
  keyboard-shortcuts.md          /shortcuts and the index of every shortcut
  the-utilities-menu.md          the header menu: Resume, Rules, Extensions, Files, Backups,
                                 Settings, Keyboard shortcuts
  language-and-theme.md          the UI language, the AI-translated Chinese option, light and dark

cross-cutting/
  persistence-and-restore.md     what is saved when, where, and what a reload brings back
  several-taskpanes.md           two workbooks, two windows on one workbook, and the workbook lock
                                 across them
  provider-errors-and-offline.md every way a provider call can fail and what the user sees
  host-differences.md            Excel desktop on macOS and Windows versus Excel on the web
```

## Coverage

Status is one of `not started`, `drafted`, or `verified`.

| Document | Status |
| --- | --- |
| glossary.md | not started |
| bug-triage.md | not started |
| verification/ (5 checklists) | not started |
| foundations/the-taskpane.md | not started |
| foundations/turns.md | not started |
| foundations/sessions-and-tabs.md | not started |
| foundations/the-workbook.md | not started |
| foundations/models-and-providers.md | not started |
| turn/the-composer.md | not started |
| turn/sending-and-the-response.md | not started |
| turn/tool-calls.md | not started |
| turn/steering-and-the-queue.md | not started |
| turn/stopping.md | not started |
| turn/thinking.md | not started |
| turn/context-and-compaction.md | not started |
| turn/auto-context.md | not started |
| sessions/renaming-a-tab.md | not started |
| sessions/tabs.md | not started |
| sessions/reopening-closed-tabs.md | not started |
| sessions/resuming-a-session.md | not started |
| sessions/export-and-copy.md | not started |
| commands/the-command-menu.md | not started |
| commands/empty-state-and-hints.md | not started |
| workbook/reading.md | not started |
| workbook/writing.md | not started |
| workbook/confirm-mode-and-approvals.md | not started |
| workbook/backups-and-revert.md | not started |
| workbook/manual-backups.md | not started |
| setup/welcome-and-login.md | not started |
| setup/the-model-picker.md | not started |
| setup/settings.md | not started |
| setup/custom-gateways.md | not started |
| setup/the-proxy.md | not started |
| guidance/rules.md | not started |
| guidance/conventions.md | not started |
| guidance/skills.md | not started |
| files/the-files-workspace.md | not started |
| extensions/the-extensions-hub.md | not started |
| extensions/web-search-and-mcp.md | not started |
| extensions/experimental.md | not started |
| ui/the-status-bar.md | not started |
| ui/message-rendering.md | not started |
| ui/keyboard-shortcuts.md | not started |
| ui/the-utilities-menu.md | not started |
| ui/language-and-theme.md | not started |
| cross-cutting/persistence-and-restore.md | not started |
| cross-cutting/several-taskpanes.md | not started |
| cross-cutting/provider-errors-and-offline.md | not started |
| cross-cutting/host-differences.md | not started |

## Reference

The source of truth is the pi-for-excel repository itself (this repo), at the commit each document cites. The relevant locations are:

- `src/taskpane.html`, `src/taskpane.ts`, `src/boot.ts`: the surface this project describes
- `src/taskpane/`: where interaction state lives (bootstrap in `init.ts`, `keyboard-shortcuts.ts`, `action-queue.ts`, `session-runtime-manager.ts`, `sessions.ts`, `tab-layout*.ts`, `recently-closed.ts`, `status-bar.ts`, `status-popovers.ts`, `welcome-login.ts`)
- `node_modules/@earendil-works/pi-agent-core`: the agent loop (streaming, tool execution, steer and follow-up queues, abort)
- `src/ui/`: the UI (`pi-sidebar.ts`, `pi-input.ts`, `working-indicator.ts`, `messages/`, `tool-renderers.ts`, `toast.ts`, `overlay-*.ts`, dialogs)
- `src/commands/`: slash commands (`builtins/`, `command-menu.ts`, `busy-command-policy.ts`)
- `src/tools/`, `src/workbook/`, `src/audit/`: the workbook tools, the workbook coordinator and lock, automatic backups and recovery
- `src/language/locales/en.json`: every user-visible string, and the source for quoting UI wording
- `tests/`: behavioral tests
- `src/compaction/`, `src/context/`, `src/models/`, `src/auth/`, `src/storage/`, `src/files/`, `src/extensions/`: subsystems that shape the experience (compaction thresholds, auto-context, model ordering, login, persistence, the Files workspace, extensions)
