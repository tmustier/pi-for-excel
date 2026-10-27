# Hand verification

The feature documents were written from the code. This directory is the protocol for checking them against the running product, one observable claim at a time.

## What is here

| File | Covers |
| --- | --- |
| [foundations-and-turn.md](foundations-and-turn.md) | `foundations/the-taskpane.md`, `foundations/turns.md` and `turn/*` |
| [sessions-and-commands.md](sessions-and-commands.md) | `foundations/sessions-and-tabs.md`, `sessions/*` and `commands/*` |
| [workbook.md](workbook.md) | `foundations/the-workbook.md` and `workbook/*` |
| [setup.md](setup.md) | `foundations/models-and-providers.md` and `setup/*` |
| [guidance-files-extensions.md](guidance-files-extensions.md) | `guidance/*`, `files/*` and `extensions/*` |
| [ui-and-cross-cutting.md](ui-and-cross-cutting.md) | `ui/*` and `cross-cutting/*` |

The agent's own experience (what the model is given and what its tools return) is checked separately in [agent-description/verification/](../../agent-description/verification/).

Each file has one table per document. Each row is an item with a stable ID (`TABS-08`, `WB-13`), a priority, what it needs (the Device column), the claim with a link to the document section, the setup, numbered steps, the expected result and a Result column. Items that cannot be checked by hand are listed under each document as "Not checkable by hand".

Priorities: **P1** is an established fact, a claim many documents depend on, or a suspected bug. **P2** is an ordinary claim. **P3** is a number, a color or a timing.

## How to run a pass

1. Bring up Pi in Excel. In an Amp orb, `node scripts/orb/excel-web-session.mjs --check` opens Pi in Excel on the web, already logged in to a model (see [docs/orb-excel-web.md](../../docs/orb-excel-web.md)). On a Mac or Windows machine, sideload `manifest.xml` and run `npm run dev` (see [docs/install.md](../../docs/install.md)). Start each file from the fresh state its first paragraph describes.
2. Confirm the commit. Every document says which commit it was verified against (`ec453df` for this set). Run `git rev-parse --short HEAD` in the source repo; if it differs, some failures will be drift, not defects.
3. Read the linked section before each item. The item is a summary; the section is the claim.
4. Work through P1 first across all files, then P2, then P3.
5. Record `pass`, `fail` or `blocked` in the Result column, with a note for anything other than a clean pass. A fail is something the document says that the product does not do. A blocked item could not be run; say what it needs.
6. File every fail in [bug-triage.md](../bug-triage.md). If the entry exists, add the item ID under "Raised by"; if not, add an entry. When the document was wrong rather than the product, correct the document and write "document corrected" in the Result note instead.
7. When every P1 and P2 item for a document has passed or been filed, change its row in the [coverage table](../README.md#coverage) from `drafted` to `verified`. Only a human pass does this.

## Devices and conditions

- **web**: Excel on the web in Chrome. Chrome takes Ctrl+T, Ctrl+W (which closes the Excel window), Ctrl+PageUp/PageDown and Ctrl+Shift+T before the taskpane sees them; `window.prompt`, `alert`, `confirm` and clipboard writes are blocked in the add-in frame; `workbook.isDirty` is always false.
- **web-anon**: a workbook opened through an anonymous sharing link, as the orb session does. Pi cannot identify it, so there is no workbook filter, no workbook rules, no file tagging, and all such workbooks share one tab layout.
- **desktop**: Excel on macOS or Windows with a workbook saved to disk. The only way to check saving, Save As, external links and desktop shortcuts.
- **2 tabs**: two tabs in one taskpane. **2 panes**: two taskpanes sharing storage; on the web, open the same workbook link in a second browser tab of the same profile. A second user (co-authoring) is a second, fresh browser profile on the same link.
- **slow write**: a mutation that holds the workbook lock long enough to observe. Delay `Excel.run` in the taskpane over the DevTools protocol (a few seconds is enough); without it the lock is held for milliseconds.
- **long session**: compaction and context items need a session with real content. Use a cheap model (GPT-6.1 Luna or GPT-6 Luna) and have Pi read a few thousand cells of text. To trip auto-compaction without hundreds of thousands of tokens, give the model a small context window (a custom gateway model, or an in-memory override).
- **reload**: reload the taskpane page (DevTools `Page.reload` on the taskpane target), not the whole Excel window, unless the item says otherwise.

## Driving the product from a console or script

In the orb the taskpane is a DevTools-protocol target (`https://localhost:3141/src/taskpane.html?et=`, port 9223). It is good for setting up exact state (clearing IndexedDB stores, delaying `Excel.run`, intercepting a model request to hold or fail it) and for reading state back (cell values, `isDirty`, the DOM of a tool card, the request sent to the model). Use real input for items about input: page-level mouse and key events through the protocol reach the shortcut layer, but browser-reserved shortcuts need real OS keys (`wtype` in the orb), and hover effects can differ from a physical mouse (MENU-15 and PICKER-08 were seen only through the protocol and are marked for a physical-mouse repeat).

Never log out of or refresh the shared login (`PI_EXCEL_WEB_AUTH_JSON`): refreshing rotates its token and logs out every other orb using it. Test OAuth refresh with a separate login.

## Results so far

First pass: 26 October 2026, Excel on the web in Amp orbs (Chrome over the DevTools protocol, signed-out anonymous-link workbook), against `ec453df`, one agent per file, with real model turns on GPT-6.1 Sol, GPT-6 Luna and GPT-6.1 Luna. Blocked items are being rerun: desktop items on a Mac, the rest in the orbs with the techniques above.

| File | Items | Pass | Fail | Blocked | Not run |
| --- | --- | --- | --- | --- | --- |
| foundations-and-turn.md | 147 | 135 | 11 | 0 | 1 |
| sessions-and-commands.md | 129 | 116 | 13 | 0 | 0 |
| workbook.md | 115 | 91 | 24 | 0 | 0 |
| setup.md | 134 | 125 | 9 | 0 | 0 |
| guidance-files-extensions.md | 156 | 142 | 14 | 0 | 0 |
| ui-and-cross-cutting.md | 118 | 111 | 3 | 3 | 1 |

### Combined fixes (`integration-fixes`)

Second pass: 27 October 2026, on the `integration-fixes` branch (all open fix PRs merged together; a2b3f2d), one independent verifier per cluster, results in [integration-fixes/](integration-fixes/).

| Cluster | Items | Pass | Fail | Blocked |
| --- | --- | --- | --- | --- |
| [setup](integration-fixes/setup.md) | 134 | 132 | 2 | 0 |
| [sessions-and-commands](integration-fixes/sessions-and-commands.md) | 129 | 129 | 0 | 0 |
| [foundations-and-turn](integration-fixes/foundations-and-turn.md) | 147 | 144 | 3 | 0 |
| [workbook](integration-fixes/workbook.md) | 115 | 115 | 0 | 0 |
| [ui-and-cross-cutting](integration-fixes/ui-and-cross-cutting.md) | 118 | 113 | 1 | 4 |

Desktop rows were run in Excel for Mac on e48d83c and 8099cd6. Still blocked: KEYS-19 (needs Windows), PERSIST-12 and HOST-09 (need a OneDrive sign-in in Excel), MULTI-10 (rerun with a separate login pending).

A fail on a "(suspected bug)" row means the bug reproduced; those are filed in [bug-triage.md](../bug-triage.md). Fails that were document errors have been corrected in the documents and are noted in the row. No document is marked `verified`: that needs a human pass.
