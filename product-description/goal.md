# Goal: complete the Pi for Excel product description

You are working in `product-description/` of the pi-for-excel repo. Read `README.md`, `glossary.md`, `foundations/turns.md`, and `sessions/renaming-a-tab.md` first. The README defines the purpose, the variant rows, the interrupt list, the cross-cutting order, the method, the structure, and the coverage table. The other three are the exemplars: match their depth, tone, and structure exactly. Your job is to write every document in the README's structure until the coverage table has no `not started` rows, then run a consistency pass.

## Source of truth

The pi-for-excel source is the rest of this repository (everything outside `product-description/`). Describe the experience in the Pi taskpane inside Excel (`src/taskpane.html`, booted by `src/taskpane.ts` and `src/taskpane/init.ts`), with one provider connected, in the default configuration: no extensions installed, no experimental features enabled, no rules or conventions set, execution mode at its default (Auto). WPS Spreadsheets, extension authoring, the local bridges and the debug UI are out of scope (see the README's scope decisions).

For each document, read in this order before writing:

1. Where the interaction state lives for the feature: `src/taskpane/init.ts` (search for the feature's handler: `sidebar.on…`, `pi:` events, the command's action), `src/taskpane/keyboard-shortcuts.ts` and `keyboard-shortcuts/`, `src/taskpane/action-queue.ts`, `src/taskpane/session-runtime-manager.ts`, and the command in `src/commands/builtins/`.
2. The shared pipeline where relevant: the agent loop in `node_modules/@earendil-works/pi-agent-core/dist/` (prompt, steer, followUp, abort, waitForIdle), `src/commands/slash-command-execution.ts` and `busy-command-policy.ts`, `src/workbook/coordinator.ts` (the workbook lock), `src/tools/with-workbook-coordinator.ts`, `src/workbook/recovery-log.ts` (checkpoints), `src/compaction/`.
3. The tests in `tests/`. They are close to executable specifications of edge cases. Key files: `action-queue.test.ts`, `session-runtime-manager.test.ts`, `session-runtime-lifecycle.test.ts`, `tab-layout.test.ts`, `tab-layout-persistence.test.ts`, `session-title.test.ts`, `auto-compaction.test.ts`, `overflow-recovery.test.ts`, `compaction-thresholds.test.ts`, `status-popovers.test.ts`, `context-injection.test.ts`, `workbook-write-inspect-undo-contract.test.ts`, `recovery-*.test.ts`, `error-path-matrix.test.ts`, and the browser tests in `tests/browser/`.
4. UI behavior: `src/ui/` (`pi-sidebar.ts` for tabs, the empty state and the utilities menu; `pi-input.ts` for the composer; `working-indicator.ts`; `messages/` and `tool-renderers.ts` for the transcript; `toast.ts`, `overlay-dialog.ts`, `overlay-escape.ts`, `confirm-dialog.ts`, `text-input-dialog.ts` for overlays).
5. Wording: `src/language/locales/en.json`. Quote UI strings exactly from here.
6. Defaults and thresholds: `src/compaction/defaults.ts`, `src/models/`, `src/execution/mode.ts`, `src/taskpane/default-model.ts`, and constants at the top of the files above.

Do not describe code. Describe what the user sees and does. Technical detail goes only in `> Technical note:` block quotes, and only when the mechanism changes what the user would expect.

## Document skeleton

Every feature document uses all eight sections, in this order, with these headings. Foundations, UI, and cross-cutting documents may drop sections that do not apply (a settings page has no busy phase) but must still cover cancel and interrupt behavior wherever an interaction exists. A small feature done properly runs 150–200 lines; the turn documents run 250–300.

The H1 is "The {feature}" for screens, controls and UI ("The composer", "The status bar"), a bare noun or command for actions and commands ("Renaming a tab", "`/backup`"), and a descriptive phrase for cross-cutting concerns ("Persistence and restore").

1. `## Summary` — one paragraph, abstract: what the feature is and what it lets the user do. Then the concrete identity: where it lives, how it is reached (a shortcut, a button, a command, a menu item), what signals that it is active, and whether it is available while the tab is busy, with no provider connected, or on Excel on the web.
2. `## The simple case` — the common path in prose, two to four short paragraphs. No variants, no edge cases. Note where the user lands afterwards.
3. `## The interaction, event by event` — a Mermaid `stateDiagram-v2` of only the states the user passes through, every transition labeled with its trigger and, in parentheses, what it commits or discards. Then five `###` subsections in this order, renamed to the feature's own words but always occupying the same five slots:
    - **Compose** — what begins the interaction and what is decided at that instant: what is targeted, what is prefilled or focused, what is validated before anything runs. Which variant, if set at the start, changes the outcome.
    - **Ends at once** — the short path: empty input, a command that resolves locally, a cancelled dialog, a blocked action. What, if anything, is recorded. Say explicitly when nothing is recorded.
    - **Sent** — what crosses the line from free-to-abandon to committed (the message joins the transcript, the first workbook change, the title is saved). What is fixed from then on. What visual begins.
    - **While busy** — what updates live, what the user can still do, what is disabled.
    - **Finishes** — what is persisted and where; what changed in the workbook and whether it can be reverted; what the feature does next; the failure path.
4. `## Modifiers` — a table with columns `Modifier | Set before | Changed while busy` and exactly these rows, in this order, every cell filled ("No effect." where that is the answer):
    - Model and provider
    - Thinking level
    - Execution mode
    - Files in the workspace, or an overlay open
    - Tab state (new, restored, another tab busy)
    - Host (desktop or web)
   Then a sentence on what changing a variant mid-interaction does.
5. `## Cancel and interrupt` — a table with columns `Event | Before sent | While busy` and exactly these rows, in this order, every cell filled:
    - Stop (button or Escape)
    - The user sends something else
    - Tab switched, created, closed or reopened
    - An overlay or menu opens
    - Model or thinking level changes
    - The provider fails
    - The taskpane goes away
    - The workbook changes underneath
    - Focus moves to the grid or another window
    - The context window fills
   Follow with a sentence or two on what happens after the interrupt (does the user stay in the feature? is a draft kept? is partial work left in the workbook?).
6. `## Interactions with other systems` — one bold-led paragraph per concern, in this order, even when the answer is "No interaction.": **Persistence.** **Workbook changes and recovery.** **Other tabs and sessions.** **Execution mode and approvals.** **Models and providers.** **Offline and provider errors.** **Notifications.** **Settings and preferences.** **Host differences.** **Language.** **Accessibility and keyboard.**
7. `## Edge cases` — a bullet list: limits, empty states, repeated invocation, ambiguous targets, the first and last instance of something.
8. `## Open questions and verification` — bullets: what was read from code but not confirmed by hand; behavior that looks like a bug, stated as such ("may be worth treating as a bug rather than documenting"); assumptions. Then the footer line: ``Verified against pi-for-excel commit `{sha}`.``

## Writing rules

- Follow the skeleton above for every feature document. The variant rows, interrupt rows and cross-cutting order are fixed; do not add, drop, or reorder them in a single document.
- Use the glossary's words. If you need a term the glossary lacks, add it to `glossary.md` in the right section with a one-paragraph definition, then use it. Never coin a synonym for an existing term.
- Sentence case for all headings. Direct, concrete language. No hedging, no marketing.
- State surprising behavior plainly and say why if the reason is in the code or a comment. If it looks like a bug, say so in "Open questions" rather than smoothing it over.
- Cross-reference other documents with relative links rather than repeating their content. [foundations/turns.md](foundations/turns.md) owns what busy, steer, follow-up, queued and Stop mean; [foundations/the-workbook.md](foundations/the-workbook.md) owns the workbook lock, checkpoints and execution mode; [foundations/sessions-and-tabs.md](foundations/sessions-and-tabs.md) owns what a session and tab are; [cross-cutting/persistence-and-restore.md](cross-cutting/persistence-and-restore.md) owns what is saved when. Do not restate them; link.
- Quote UI wording exactly from `en.json`, in double quotes.
- Every document ends with "## Open questions and verification" followed by ``Verified against pi-for-excel commit `{sha}`.`` using the source commit below.
- Mermaid `stateDiagram-v2` for each interaction's states. Keep it to the states the user passes through.

## Things already established (do not re-derive, do not contradict)

- The source commit for this pass is `ec453df`. Commits that only touch `product-description/` do not change behavior; keep citing `ec453df` unless the source moves.

## Order of work

1. `sessions/renaming-a-tab.md` (pilot), then `foundations/` in this order: `the-taskpane.md`, `turns.md`, `sessions-and-tabs.md`, `the-workbook.md`, `models-and-providers.md`. Everything else links to them.
2. `turn/` next, all eight documents. This is the hardest part and the bulk of the experience. Read `keyboard-shortcuts.ts`, `keyboard-shortcuts/editor-actions.ts`, `action-queue.ts`, `queue-display.ts`, `pi-input.ts`, `working-indicator.ts`, and the agent loop before starting any of them, because the states hand off to each other and the documents must agree on where one ends and the next begins.
3. The remaining areas: `sessions/`, `commands/`, `workbook/`, `setup/`, `guidance/`, `files/`, `extensions/`, `ui/`, `cross-cutting/`. These are independent of each other and can be drafted in parallel with subagents once the foundations and turn documents exist to link to. If you parallelize, give each subagent this file, the exemplars, and the specific documents to write; then review every result yourself for consistency with the glossary and the established facts above before accepting it.
4. Consistency pass over the whole set: same term for the same thing everywhere, no two documents describing the same behavior differently, every relative link resolves, every document has a verification footer, every glossary term used is defined.
5. Update the coverage table in `README.md` as you go: `drafted` when written, never `verified` (verification by hand is a separate pass).

## Working rules

- Commit after each document or coherent group of documents with a message of the form `docs: add {path}` or `docs: revise {path}`. This repo does not add AI attribution to commit messages.
- Do not modify anything outside `product-description/`. The rest of the repo is read-only reference material.
- Do not add files outside the README's structure without updating the structure and coverage table to match.
- When a behavior cannot be determined from code and tests, write down what you could determine, put the rest in "Open questions", and move on. Do not guess and do not block.
- Depth bar: `sessions/renaming-a-tab.md` is the reference for a small feature. The turn documents will be longer; UI documents will often be shorter. Completeness matters more than length. Every state, every modifier row, every interrupt row must be accounted for, even if the answer is "No effect."
- If the README's structure is wrong for something you discover (a document that should be split, two that should merge), make the change, update the structure and coverage table, and note why in the commit message.

You are done when the coverage table has no `not started` rows, the consistency pass is complete, and everything is committed.
