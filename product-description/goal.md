# Goal: maintain the Pi for Excel product description

Read [README.md](README.md), [glossary.md](glossary.md), [turns](foundations/turns.md) and [renaming a tab](sessions/renaming-a-tab.md) first. The README defines the purpose, vocabulary, scope, structure and coverage table. Describe the experience from outside in, not the implementation.

## Source of truth

The source baseline is **main `1f51cf8`**, including the merged fixes below. Documentation-only commits do not change that baseline. Read the owning code, tests and UI wording in `src/language/locales/en.json`; check ambiguous behavior in the running product. Old verification runs are historical evidence, not the current contract. Never label a behavior confirmed solely because a PR description claims it.

The default configuration is one connected provider, Auto execution mode, no installed extensions, experimental flags, rules or conventions. WPS and extension authoring are out of scope. The agent-facing counterpart is [agent-description](../agent-description/README.md).

## Document skeleton

Every feature document follows the README's eight sections:

1. **Summary:** what the feature does, where it lives, availability while busy and on the web.
2. **The simple case:** the common path and where the user lands.
3. **The interaction, event by event:** compose, ends at once, sent, while busy, finishes. Include a small Mermaid `stateDiagram-v2`, with triggers on transitions and what is committed or discarded.
4. **Modifiers:** `Modifier | Set before | Changed while busy`, in this order: model and provider; thinking level; execution mode; Files or overlay; tab state; host. Fill every cell, including "No effect."
5. **Cancel and interrupt:** `Event | Before sent | While busy`, in this order: Stop; another send; tabs; overlay/menu; model/thinking; provider failure; taskpane goes away; workbook changes; focus moves; context fills. Say what remains after an interrupt.
6. **Interactions with other systems:** bold-led paragraphs in this order: Persistence; Workbook changes and recovery; Other tabs and sessions; Execution mode and approvals; Models and providers; Offline and provider errors; Notifications; Settings and preferences; Host differences; Language; Accessibility and keyboard.
7. **Edge cases:** limits, repeated invocation, empty states, ambiguous targets.
8. **Open questions and verification:** distinguish code inspection from live observation, retain genuine unknowns, and end with `Verified against pi-for-excel commit` and the source baseline.

Foundations and cross-cutting documents may omit phases that do not apply. Keep existing heading anchors when revising: other documents and historical checklists link to them.

## Writing rules

- Sentence case, direct prose, glossary terms and exact UI wording in quotes.
- Technical detail belongs only in `> Technical note:` blocks and only where it changes expectations.
- Link to the owner of a shared behavior instead of duplicating it. Foundations own turns, sessions and the workbook; [persistence and restore](cross-cutting/persistence-and-restore.md) owns storage.
- State surprising behavior plainly. Keep unresolved defects in Open questions; remove resolved bug claims from the current contract, not from historical observations.
- Do not turn current product-review decisions into shipped behavior. Their implementation PRs own that documentation.
- Do not mark a document verified just because it was revised. Coverage remains `drafted` until its required acceptance items pass or are filed.

## Things already established

- A turn spans the user's send through idle, possibly several model responses and tool batches. Steers wait for the current response and tools; follow-ups wait until Pi would otherwise finish.
- Each tab owns its composer draft, Not sent rows, transcript, model, queue, activity and error banner. Drafts and Not sent rows are saved with the session; they are never sent automatically on restore.
- Stop button, Escape when not claimed by a menu/modal, and Stop and close use one stop path. It cancels retry waits and compaction, holds undelivered messages as separate Not sent rows, and names cancelled commands in a toast. Work already applied in Excel stays applied.
- A session saves from its first user message. Chosen titles, stable tab numbers, recently closed tabs and identity-less workbook tab layouts survive reload. Resume/reopen focuses an already-open session rather than duplicating it.
- Execution mode defaults to Auto and is shared across taskpanes on this browser origin. Confirm asks once before each workbook mutation, before the lock. Approvals queue, name the tab and action, and are withdrawn by Stop. Deleting recovery backups always asks; local Files/extension consent follows the rule in `src/tools/DECISIONS.md`.
- The workbook write lock serializes mutations across tabs in a taskpane; it is not a co-authoring lock. Same-runtime reads wait for earlier registered mutations, not mutations in other runtimes.
- Saving or reloading never clears automatic backups. Retention applies per workbook, with a global cap of 240. A restore makes a rollback backup when possible and reports failures to save it.
- Context meter warnings use the selected model's compaction budget, not fixed percentages of its full window. Compaction belongs to its originating tab and saves immediately. Summaries distinguish answered requests and shrink large answered results in the kept tail.
- Workbook structure, Files summary and other people's edits are stored conversation notes when changed. Selection is fresh per request with a bounded sample. Tools wait for cell-edit mode for up to 60 seconds before a precise failure; auto-context does not hold up a request for that wait.
- Large reads page; compact/detailed reads use displayed values and annotate hidden stored values and workbook state. Checked writes refuse merge/spill/protection obstacles, chunk large inputs, and report the scope and limitations of read-back verification.
- Every modal takes focus, traps Tab and returns focus; Escape acts on the topmost surface. Shortcuts do not operate behind a modal. Shift+Tab cycles thinking only in the composer.
- The notice area is below the tabs. Banners are dismissible; a disabled proxy produces no warning. Excel-web loopback restrictions are explained rather than offering an impossible permission prompt.
- Clipboard copy uses fallbacks; bare `/export` downloads when copy is unavailable. Downloads use a single blob download link without opening a spare browser tab. Desktop Save-sheet behavior needs desktop acceptance.
- Rules, conventions, skills and integration changes refresh runtime capabilities without a reload. Conflicting Rules saves ask before replacement; same-name imports receive a new name rather than overwriting.

## Merged fixes

PRs [#786](https://github.com/tmustier/pi-for-excel/pull/786)–[#799](https://github.com/tmustier/pi-for-excel/pull/799) merged into main at `1f51cf8` on 27 October 2026:

- #786 transcript rendering and details mode; #787 paged reads and workbook-state reporting.
- #788 Rules, conventions, skills and Files refresh/safety; #789 per-tab turn state, Stop, drafts and retries.
- #790 integration permissions and local-change consent; #791 checked, chunked writes and truthful results.
- #792 durable auto-context, bounded selection and Excel waits; #793 shared overlay, menu, notice and status-bar behavior.
- #794 durable backups, queued approvals and shared mode; #795 sessions, stable tabs, model selection and copy fallbacks.
- #796 host-specific download, proxy, language and startup handling; #797 Python backend failure and checked transform behavior.
- #798 answered-request-aware compaction; #799 bundled Copilot device-code login and cancellation.

This is a provenance record, not a substitute for reading the merged code or for product verification.

## Order of work

Revise foundations and shared contracts first, then their feature documents and checklist expectations. Keep recorded verification results intact and identify their older source baseline. Run the relative-link and heading-anchor checker before each push:

```sh
python3 -I product-description/scripts/check-links.py product-description
```

The checker is committed with its source credited. Check `agent-description` as well when changing its links.

## Working rules

- Keep changes within `product-description/` and `agent-description/` for this revision.
- Do not edit `product-description/bug-triage.md`, `product-description/review/README.md` or either verification README owned by the coordinating pass.
- Pull with rebase before every commit. Commit small coherent groups with `docs: revise <file> to merged behaviour`, no attribution, and `--no-verify`. Push only to `product-description`; do not merge or mark the draft PR ready.
- Unknown behavior belongs in Open questions, not an invented rule. Report contradictions with a merged PR with reproducible evidence.
