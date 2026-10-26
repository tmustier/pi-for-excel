# Verification: guidance, files and extensions

How to run this file: bring up Pi in Excel on the web with `node scripts/orb/excel-web-session.mjs --check` (see [docs/orb-excel-web.md](../../docs/orb-excel-web.md)). The workbook is opened signed out through the shared folder link, so it has no workbook identity: "This file" rules, the files "Workbook" badge and the workbook web-search scope behave as for an unidentified workbook. Never log out of, disconnect or refresh the shared OpenAI (ChatGPT) credential. Use made-up API keys and an unreachable MCP URL only. Keep model turns few and tiny; switch to GPT-6 Luna for anything that needs more than a couple. Before each section close Settings and Files, keep execution mode on Auto unless the item says Confirm, and afterwards clear any key, MCP server, external skill and test file you added.

Device column: `web` is the Excel on the web session; `desktop` is Excel desktop (not available in an orb; those items go to a desktop pass); `second taskpane` needs two taskpanes open at once.

## guidance/rules.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| RULES-01 | P1 | web | `/rules` opens the Settings overlay on "Rules", tab "All my files", focus in the text box ([opening the editor](../guidance/rules.md#opening-the-editor)). | Idle tab, overlay closed. | 1. Type `/rules`, Enter. | Rules page, "All my files" selected, focus in the textarea. | pass |
| RULES-02 | P2 | web | The hint, placeholder and "0 / 2,000 chars" counter ([the simple case](../guidance/rules.md#the-simple-case)). | No user rules. | 1. Open `/rules`. | Hint "Guidance given to Pi in all your conversations…", placeholder "Your preferences and habits…", "0 / 2,000 chars". | pass |
| RULES-03 | P2 | web | The utilities-menu item "Rules" opens the same page ([opening the editor](../guidance/rules.md#opening-the-editor)). | — | 1. Gear → "Rules". | Rules page. | pass |
| RULES-04 | P2 | web | The Settings row "Rules & conventions" opens the same page ([opening the editor](../guidance/rules.md#opening-the-editor)). | Settings root open. | 1. Click "Rules & conventions". | Rules page. | pass |
| RULES-05 | P2 | web | "This file" on an unidentified workbook shows "Workbook: Current workbook" and the "Can't identify…" hint; the box accepts typing, counter out of 4,000 ([opening the editor](../guidance/rules.md#opening-the-editor)). | Signed-out link workbook (identity unknown). | 1. Click "This file".<br>2. Type text. | Both lines shown; text accepted; "17 / 4,000 chars". | pass |
| RULES-06 | P1 | web | Save shows "Rules saved" and goes back to the "Settings" page instead of closing (suspected bug) ([saving](../guidance/rules.md#saving)). | Text typed. | 1. Click "Save". | Toast "Rules saved"; Settings root shown. Record what happens. | pass (confirmed) |
| RULES-07 | P2 | web | Cancel with no changes goes back silently ([ends at once](../guidance/rules.md#ends-at-once)). | No changes. | 1. Click "Cancel". | Settings root; no toast, no dialog. | pass |
| RULES-08 | P1 | web | The first Escape in the text box only leaves it ([ends at once](../guidance/rules.md#ends-at-once)). | Typed text, focus in box. | 1. Escape. | Page still open. | pass |
| RULES-09 | P1 | web | Workbook rules typed on an unidentified workbook are dropped on Save while "Rules saved" shows (suspected bug) ([edge cases](../guidance/rules.md#edge-cases)). | Unidentified workbook. | 1. Type under "This file".<br>2. Save.<br>3. Reopen, "This file". | "Rules saved"; box empty on reopen. Record what happens. | pass (confirmed) |
| RULES-10 | P1 | web | Escape with unsaved changes asks "Discard changes?" ([ends at once](../guidance/rules.md#ends-at-once)). | Typed text, focus out of box. | 1. Escape. | "Discard changes?", "You have unsaved changes to your rules. Discard them?", "Keep editing", "Discard". | pass |
| RULES-11 | P1 | web | Escape on the question keeps editing with drafts intact ([ends at once](../guidance/rules.md#ends-at-once)). | Question open. | 1. Escape. | Question closes; counter unchanged. | pass |
| RULES-12 | P2 | web | × with changes asks; "Discard" closes the overlay ([ends at once](../guidance/rules.md#ends-at-once)). | Typed text. | 1. Click ×.<br>2. "Discard". | Question, then overlay closed. | pass |
| RULES-13 | P1 | web | An unsaved change on "Formats" triggers the question from "All my files" ([ends at once](../guidance/rules.md#ends-at-once)). | Change a convention, switch to "All my files". | 1. Escape. | "Discard changes?". | pass |
| RULES-14 | P3 | web | Counter warns past 2,000; Save is still allowed ([opening the editor](../guidance/rules.md#opening-the-editor)). | — | 1. Enter 2,050 characters.<br>2. Save. | "2,050 / 2,000 chars" in warning style; saved in full. | pass |
| RULES-15 | P2 | web | A blank (whitespace) box saved clears the rule ([saving](../guidance/rules.md#saving)). | User rules set. | 1. Replace with spaces.<br>2. Save. | Stored value empty. | pass |
| RULES-16 | P2 | web | Pi's `instructions` tool card reads "Remember user rules" with "Updated user rules (N/2000 chars):" ([when pi changes the rules](../guidance/rules.md#when-pi-changes-the-rules)). | Auto mode. | 1. Ask Pi to append a user rule. | Card and result as described. | pass |
| RULES-17 | P2 | web | No toast when Pi changes rules ([when pi changes the rules](../guidance/rules.md#when-pi-changes-the-rules)). | As RULES-16. | 1. Watch toasts. | No toast. | pass |
| RULES-18 | P1 | web | The `instructions` tool asks no approval in Confirm mode ([when pi changes the rules](../guidance/rules.md#when-pi-changes-the-rules)). | Confirm mode. | 1. Ask Pi to add a rule. | No dialog. | blocked: not run (turn budget) |
| RULES-19 | P1 | web | On an unidentified workbook the tool fails with "Error: workbook identity unavailable…" ([when pi changes the rules](../guidance/rules.md#when-pi-changes-the-rules)). | Unidentified workbook. | 1. Ask Pi to set workbook rules. | Error in the card. | blocked: not run (turn budget) |
| RULES-20 | P2 | web | Rules are stored in add-in storage, not the workbook ([after saving](../guidance/rules.md#after-saving)). | — | 1. Save a rule.<br>2. Read IndexedDB `settings` → `user.instructions`. | Value present in storage. | pass |
| RULES-21 | P1 | web | `/rules` opens while the tab is busy ([summary](../guidance/rules.md#summary)). | Tab responding. | 1. Type `/rules`, Enter. | Page opens. | pass (opened about 2 s into a short turn) |
| RULES-22 | P2 | web | Escape on the page never stops a responding tab ([cancel and interrupt](../guidance/rules.md#cancel-and-interrupt)). | Long turn running, page open. | 1. Escape twice. | Turn continues. | blocked: no turn lasted long enough |
| RULES-23 | P1 | web | Unsaved drafts are lost on reload with no warning ([cancel and interrupt](../guidance/rules.md#cancel-and-interrupt)). | Typed text. | 1. Reload the taskpane. | Draft gone, no prompt. | blocked: reload not risked (repeats credential restore) |
| RULES-24 | P1 | web | Saving the page overwrites a rule Pi saved while the page was open (suspected bug) ([edge cases](../guidance/rules.md#edge-cases)). | Page open on empty rules. | 1. Have Pi append "Prefer GBP".<br>2. Save the page unchanged.<br>3. Read storage. | Page still empty; stored rule overwritten. Record what happens. | pass (confirmed) |
| RULES-25 | P2 | desktop | Saving a new workbook for the first time while the page is open leaves "This file" unusable until reopened ([cancel and interrupt](../guidance/rules.md#cancel-and-interrupt)). | Never-saved workbook. | 1. Open page.<br>2. Save workbook.<br>3. Check "This file". | Still the unidentified hint. | pass. desktop (macOS): on a new workbook the page said "Can't identify this workbook right now — try saving the file first."; after Save As it still did until reopened, when it showed "Workbook: pd-desktop-C.xlsx". |
| RULES-26 | P2 | second taskpane | A second taskpane picks up a change at its next refresh ([interactions with other systems](../guidance/rules.md#interactions-with-other-systems)). | Two taskpanes. | 1. Save in one.<br>2. Focus the other. | — | blocked. desktop (macOS): two windows on one workbook each open their own Pi (Window > New Window, then "Open Pi"), so the setup exists; but a rule saved in window 1 was reported as "NONE" by the model in both windows, including the one that saved it, so a model turn cannot show whether the other pane picked it up, and the pane shows no rules state. |

Not checkable by hand:

- Whether a save reaches every open tab's next model request (visible only in the request).
- Whether Pi follows contradicting user and workbook rules by asking ([Edge cases](../guidance/rules.md#edge-cases)).

## guidance/conventions.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| CONV-01 | P2 | web | The Formats tab hint and the six built-in cards with the default strings ([what the tab shows](../guidance/conventions.md#what-the-tab-shows)). | Default conventions. | 1. `/rules` → "Formats". | Strings as in [The defaults](../guidance/conventions.md#the-defaults). | pass |
| CONV-02 | P2 | web | Expanded Currency shows Format, three previews, dp/neg/zero/symbol ([editing a format](../guidance/conventions.md#editing-a-format)). | — | 1. Expand "Currency". | As described. | pass |
| CONV-03 | P1 | web | A quick-option change rebuilds the string and previews at once ([editing a format](../guidance/conventions.md#editing-a-format)). | Currency open. | 1. dp → 0.<br>2. symbol → €. | `€* #,##0_);€* (#,##0);€* --_)`; previews "€1,235", "€(1,235)", "€--". | pass |
| CONV-04 | P1 | web | After Save, Pi's "currency" formatting applies exactly the convention's string ([what the user sees in pis formatting](../guidance/conventions.md#what-the-user-sees-in-pis-formatting)). | CONV-03 saved; A1:A3 = 1234.5, −1234.5, 0. | 1. Ask Pi to format A1:A3 with its currency style.<br>2. Read numberFormat. | Excel numberFormat equals the saved string; cells "€1,235", "€(1,235)", "€--". | pass |
| CONV-05 | P3 | web | Default Currency previews ([editing a format](../guidance/conventions.md#editing-a-format)). | Defaults. | 1. Expand Currency. | "$1,234.50", "$(1,234.50)", "$--". | pass |
| CONV-06 | P1 | web | The Percent preview does not multiply by 100 (suspected bug) ([edge cases](../guidance/conventions.md#edge-cases)). | — | 1. Expand Percent.<br>2. Format a cell with 1234.5 in that string. | Preview "1,234.5%"; Excel "123,450.0%". Record what happens. | pass (confirmed) |
| CONV-07 | P1 | web | Excel accepts the Ratio format's bare `x` (open question) ([open questions and verification](../guidance/conventions.md#open-questions-and-verification)). | — | 1. Apply `#,##0.0x_);(#,##0.0x);--_x_)` to 1234.5, −1234.5, 0. | Record what Excel shows. | pass: "1,234.5x", "(1,234.5x)", "--" on the web; desktop not tried |
| CONV-08 | P2 | web | Typing in Format removes the quick options; previews read "Custom" ([editing a format](../guidance/conventions.md#editing-a-format)). | Number open. | 1. Type `0.000`, Tab. | Quick options gone; "Custom" previews; reset link shown. | pass |
| CONV-09 | P2 | web | The reset link restores the built-in defaults ([editing a format](../guidance/conventions.md#editing-a-format)). | CONV-08. | 1. Click "Custom format — use quick options to reset". | Number back to `#,##0.00_);(#,##0.00);--_)`. | pass |
| CONV-10 | P2 | web | "Add custom format" adds "custom-1" with Name, Description, Format and "Remove preset" ([editing a format](../guidance/conventions.md#editing-a-format)). | — | 1. Click "Add custom format".<br>2. Expand it. | As described. | pass |
| CONV-11 | P2 | web | An invalid colour is ignored without a message ([editing a format](../guidance/conventions.md#editing-a-format)). | — | 1. Type `notacolor` in "Hardcoded values", Tab. | Field shows `#0000FF` again; no toast. | pass |
| CONV-12 | P2 | web | Leaving after Formats changes asks; "Discard" saves nothing ([ends at once](../guidance/conventions.md#ends-at-once)). | CONV-10 unsaved. | 1. Cancel.<br>2. "Discard".<br>3. Read storage. | No custom preset stored. | pass |
| CONV-13 | P2 | web | Cells formatted earlier keep their format after a change ([after saving](../guidance/conventions.md#after-saving)). | Formatted cells. | 1. Change Currency; Save.<br>2. Read the cells. | Unchanged. | blocked: not run (turn budget) |
| CONV-14 | P1 | web | The `conventions` tool never asks approval in Confirm mode ([when pi changes the conventions](../guidance/conventions.md#when-pi-changes-the-conventions)). | Confirm mode. | 1. Ask Pi to change a convention. | No dialog; card "Update conventions". | blocked: not run (turn budget) |
| CONV-15 | P3 | web | Renaming a custom format to an existing name is ignored ([edge cases](../guidance/conventions.md#edge-cases)). | Two custom presets. | 1. Rename one to the other's name. | Old name kept, no message. | blocked: not run (turn budget) |

Not checkable by hand:

- The `pi:conventions-updated` event has no listener, so Pi's instructions keep stale overrides ([Open questions](../guidance/conventions.md#open-questions-and-verification)). Visible only in the request.
- Whether models follow the font-colour and default-font guidance.

## guidance/skills.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| SKILLS-01 | P2 | web | `/skills` opens "Skills" with "5 skills active (5 bundled, 0 external)" ([the simple case](../guidance/skills.md#the-simple-case)). | Defaults. | 1. `/skills`. | Status line as stated. | pass |
| SKILLS-02 | P2 | web | Five bundled skills with descriptions and "Bundled" badges ([the simple case](../guidance/skills.md#the-simple-case)). | — | 1. Read the list. | extending-pi, mcp-gateway, python-bridge, tmux-bridge, web-search. | pass |
| SKILLS-03 | P2 | web | Empty external list text ([the simple case](../guidance/skills.md#the-simple-case)). | No external skills. | 1. Read. | "No external skills installed." / "Paste a SKILL.md below to add one." | pass |
| SKILLS-04 | P2 | web | "Install skill" with nothing pasted ([ends at once](../guidance/skills.md#ends-at-once)). | — | 1. Click "Install skill". | "Paste a SKILL.md document first." | pass |
| SKILLS-05 | P2 | web | Invalid frontmatter toast ([acting on a skill](../guidance/skills.md#acting-on-a-skill)). | Paste `hello`. | 1. Install. | "Extensions: Invalid SKILL.md document: expected frontmatter with name and description." | pass |
| SKILLS-06 | P2 | web | Path separator toast ([edge cases](../guidance/skills.md#edge-cases)). | Name `bad/name`. | 1. Install. | "Extensions: Skill name cannot contain path separators." | pass |
| SKILLS-07 | P1 | web | A valid install saves, toasts, clears the box and counts ([acting on a skill](../guidance/skills.md#acting-on-a-skill)). | Valid `pd-test`. | 1. Install. | "Saved external skill: pd-test"; box empty; "6 skills active (5 bundled, 1 external)". | pass |
| SKILLS-08 | P1 | web | Switching off saves at once and updates the count ([acting on a skill](../guidance/skills.md#acting-on-a-skill)). | — | 1. Switch web-search off. | "Disabled skill: web-search"; "5 skills active (4 bundled, 1 external)". | pass |
| SKILLS-09 | P2 | web | Switching on ([acting on a skill](../guidance/skills.md#acting-on-a-skill)). | SKILLS-08. | 1. Switch on. | "Enabled skill: web-search". | pass |
| SKILLS-10 | P2 | web | Expanded external card shows Location and Remove ([acting on a skill](../guidance/skills.md#acting-on-a-skill)). | pd-test installed. | 1. Click the card. | "Location" `skills/external/pd-test/SKILL.md`, "Remove". | pass |
| SKILLS-11 | P1 | web | Remove deletes at once with no confirmation ([acting on a skill](../guidance/skills.md#acting-on-a-skill)). | SKILLS-10. | 1. Remove. | "Removed external skill: pd-test"; no dialog. | pass |
| SKILLS-12 | P2 | web | A reinstalled skill inherits its stored off state ([edge cases](../guidance/skills.md#edge-cases)). | pd-test switched off, removed. | 1. Install it again. | Its switch is off; not counted. | pass |
| SKILLS-13 | P2 | web | An external skill named like a bundled one is "Shadowed" with a disabled switch ([external skills and the files workspace](../guidance/skills.md#external-skills-and-the-files-workspace)). | — | 1. Install `name: web-search`. | "Shadowed" badge; switch disabled. | pass |
| SKILLS-14 | P2 | web | External skills appear in Files under "SKILLS" ([external skills and the files workspace](../guidance/skills.md#external-skills-and-the-files-workspace)). | pd-test installed. | 1. Open Files. | SKILLS → pd-test → SKILL.md. | pass |
| SKILLS-15 | P2 | web | Escape goes back to "Settings", then closes ([the simple case](../guidance/skills.md#the-simple-case)). | Skills page. | 1. Escape twice. | Root, then closed. | pass |
| SKILLS-16 | P2 | web | Pi's list card reads "List skills" ([when pi uses skills](../guidance/skills.md#when-pi-uses-skills)). | Auto. | 1. Ask Pi to list skills. | Card "List skills". | pass |
| SKILLS-17 | P2 | web | A read card reads "Read skill {name} (bundled)", then "(cached)" ([when pi uses skills](../guidance/skills.md#when-pi-uses-skills)). | Auto. | 1. Ask Pi to read a skill twice. | As described. | blocked: the only read was declined (card showed "Read skill web-search" without a suffix) |
| SKILLS-18 | P1 | web | A switched-off skill fails to read with "Skill not found: …" ([acting on a skill](../guidance/skills.md#acting-on-a-skill)). | Skill off. | 1. Ask Pi to read it. | Error in the card. | blocked: not run (turn budget) |
| SKILLS-19 | P2 | web | Pi installs a skill without asking in Auto ([when pi uses skills](../guidance/skills.md#when-pi-uses-skills)). | Auto. | 1. Ask Pi to install one. | Card "Install skill {name}"; no dialog. | blocked: not run (turn budget) |
| SKILLS-20 | P1 | web | In Confirm mode `skills` list asks for a workbook-mutation approval (suspected bug) ([when pi uses skills](../guidance/skills.md#when-pi-uses-skills)). | Confirm mode. | 1. Ask Pi to list skills. | Record what happens. | pass (confirmed: "Allow workbook mutation in Safe mode?", "Tool: skills:list") |
| SKILLS-21 | P1 | web | In Confirm mode `skills` read asks; Escape declines it (suspected bug) ([when pi uses skills](../guidance/skills.md#when-pi-uses-skills)). | Confirm mode. | 1. Trigger a read.<br>2. Escape. | Record what happens. | pass (confirmed: "Tool: skills:read"; result "Mutation cancelled by user (Confirm mode).") |
| SKILLS-22 | P2 | web | A file uploaded to `skills/{name}/SKILL.md` adds a skill ([external skills and the files workspace](../guidance/skills.md#external-skills-and-the-files-workspace)). | — | 1. Have Pi write the file.<br>2. Open Skills. | Listed as External. | blocked: not run (turn budget) |

Not checkable by hand:

- Whether Pi decides to read the right skill is the model's choice.

## files/the-files-workspace.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| FILES-01 | P2 | web | "Browse files" opens "Files" / "Documents available to Pi" ([the simple case](../files/the-files-workspace.md#the-simple-case)). | — | 1. Click the composer's Files button. | Overlay with Upload and Connect folder. | pass |
| FILES-02 | P2 | web | Utilities menu "Files" opens it ([summary](../files/the-files-workspace.md#summary)). | — | 1. Gear → Files. | Overlay. | pass |
| FILES-03 | P1 | web | The overlay does not take focus; the first Escape leaves the composer, the second closes; focus returns to the composer ([the simple case](../files/the-files-workspace.md#the-simple-case)). | Focus in composer. | 1. Open.<br>2. Escape.<br>3. Escape. | Open, then focus on the tab, then closed with focus in the composer. | pass |
| FILES-04 | P1 | web | `/files` while the overlay is open closes it ([summary](../files/the-files-workspace.md#summary)). | Files open. | 1. Type `/files`, Enter. | Closed. | pass |
| FILES-05 | P1 | web | Dragging files over the composer shows the drop hint ([the simple case](../files/the-files-workspace.md#the-simple-case)). | — | 1. Drag a file over the composer. | "Drop files to import to file workspace". | pass (synthetic drag events in the taskpane frame) |
| FILES-06 | P1 | web | A drop imports, toasts, and leaves the draft untouched ([importing](../files/the-files-workspace.md#importing)). | Draft "draft text". | 1. Drop `budget.csv`. | "Imported 1 file into Files."; draft unchanged; nothing sent. | pass (synthetic) |
| FILES-07 | P2 | web | Two files dropped ([importing](../files/the-files-workspace.md#importing)). | — | 1. Drop two files. | "Imported 2 files into Files." | pass |
| FILES-08 | P2 | web | Upload opens a multi-file picker and toasts ([importing](../files/the-files-workspace.md#importing)). | — | 1. Upload, choose two files. | "Imported 2 files." | pass |
| FILES-09 | P2 | web | Row lines "{size} · Uploaded · just now", newest first ([opening the list](../files/the-files-workspace.md#opening-the-list)). | Imports done. | 1. Read YOUR FILES. | As described. | pass |
| FILES-10 | P1 | web | Imported files carry a "Workbook" badge ([importing](../files/the-files-workspace.md#importing)). | Signed-out link workbook. | 1. Import.<br>2. Look for the badge. | Badge shown. | fail (document was wrong: no badge when the workbook is unidentified; corrected) |
| FILES-11 | P2 | web | BUILT-IN DOCS (47) marked "Read only" / "Pi documentation" ([opening the list](../files/the-files-workspace.md#opening-the-list)). | — | 1. Scroll. | As described. | pass |
| FILES-12 | P2 | web | SKILLS section with a folder row ([opening the list](../files/the-files-workspace.md#opening-the-list)). | pd-test installed. | 1. Read. | SKILLS 1 → pd-test → SKILL.md. | pass |
| FILES-13 | P3 | web | The footer count includes built-in docs ([edge cases](../files/the-files-workspace.md#edge-cases)). | One user file. | 1. Read the footer. | "48 files · 356 KB · Sandboxed workspace". | pass |
| FILES-14 | P1 | web | Importing a same-named file replaces it without asking (suspected bug) ([importing](../files/the-files-workspace.md#importing)). | `budget.csv` present. | 1. Drop a different `budget.csv`.<br>2. Preview. | Record what happens. | pass (confirmed: new content, no question) |
| FILES-15 | P2 | web | The filter appears at five files and matches case-insensitively ([opening the list](../files/the-files-workspace.md#opening-the-list)). | ≥5 files. | 1. Type `BUD`. | Only budget.csv. | pass |
| FILES-16 | P2 | web | No-match message ([opening the list](../files/the-files-workspace.md#opening-the-list)). | — | 1. Type `zzzq`. | "No matching files" / "Try a different filter term." | pass |
| FILES-17 | P2 | web | The first Escape in the filter only leaves the box ([cancel and interrupt](../files/the-files-workspace.md#cancel-and-interrupt)). | Focus in filter. | 1. Escape. | Overlay still open. | pass |
| FILES-18 | P2 | web | Preview: title, type line, rendered markdown, four actions ([the preview](../files/the-files-workspace.md#the-preview)). | notes.md. | 1. Click it. | "Text · 28 B · Uploaded · just now"; heading rendered; Open ↗, Download, Rename, Delete. | pass |
| FILES-19 | P3 | web | "(preview truncated)" is added for long files ([the preview](../files/the-files-workspace.md#the-preview)). | — | 1. Preview `compaction.md` (147 lines). | Suffix shown. | fail (document was wrong: the suffix needs a size-truncated read; corrected) |
| FILES-20 | P1 | web | Escape in the preview closes the whole overlay (suspected bug) ([open questions and verification](../files/the-files-workspace.md#open-questions-and-verification)). | Preview open. | 1. Escape (twice if focus is in the composer). | Record what happens. | pass (confirmed: closed, not back to the list) |
| FILES-21 | P2 | web | Built-in doc callout and actions ([the preview](../files/the-files-workspace.md#the-preview)). | — | 1. Preview a built-in doc. | "Built-in documentation — read only. …"; Copy content, Download. | pass |
| FILES-22 | P1 | web | "Copy content" toasts "Copied {name} to clipboard." ([the preview](../files/the-files-workspace.md#the-preview)). | Built-in doc open. | 1. Copy content. | Copied toast. | fail (on the web: "Copy failed: … Clipboard API has been blocked…"; document corrected) |
| FILES-23 | P2 | web | "Open ↗" opens a new window ([the preview](../files/the-files-workspace.md#the-preview)). | User file open. | 1. Open ↗. | A blob window opens. | pass |
| FILES-24 | P2 | desktop | "Download" saves the file ([the preview](../files/the-files-workspace.md#the-preview)). | User file open. | 1. Download. | File saved. | fail. desktop (macOS): "Download" and "Open ↗" in the preview of a Files item do nothing: no file, no window, no browser tab (`src/ui/files-dialog-actions.ts:70-120`). Bug. |
| FILES-25 | P1 | web | "Connect folder" connects a folder (where shown) ([importing](../files/the-files-workspace.md#importing)). | Chromium web. | 1. Connect folder. | Folder picker. | fail (product bug: "Connect folder failed: … Cross origin sub frames aren't allowed to show a file picker.") |
| FILES-26 | P2 | web | Rename dialog text, prefill and focus ([the preview](../files/the-files-workspace.md#the-preview)). | notes.md open. | 1. Rename. | "Rename file", "notes.md — leave off the extension to keep it.", prefilled, focused. | pass |
| FILES-27 | P2 | web | Rename into a folder ([the preview](../files/the-files-workspace.md#the-preview)). | Dialog open. | 1. Enter `docs/renamed`, Rename. | "Renamed to docs/renamed.md."; preview title renamed.md. | pass |
| FILES-28 | P2 | web | Delete asks and returns to the list ([the preview](../files/the-files-workspace.md#the-preview)). | renamed.md open. | 1. Delete.<br>2. Delete. | "Delete file?" with path; "Deleted renamed.md."; list. | pass |
| FILES-29 | P2 | web | ← returns to the list ([the interaction event by event](../files/the-files-workspace.md#the-interaction-event-by-event)). | Preview open. | 1. Click ←. | List. | pass |
| FILES-30 | P2 | web | A drop during a turn imports without interrupting it ([cancel and interrupt](../files/the-files-workspace.md#cancel-and-interrupt)). | Tab responding. | 1. Drop a file. | Toast; turn continues. | pass |
| FILES-31 | P2 | web | Dropping outside the composer does nothing ([importing](../files/the-files-workspace.md#importing)). | — | 1. Drop on the transcript. | No import. | blocked: not run (turn budget) |
| FILES-32 | P1 | web | Pi's `files` writes and deletes never ask, in Confirm ([what pi can do with files](../files/the-files-workspace.md#what-pi-can-do-with-files)). | Confirm mode. | 1. Ask Pi to write then delete a file. | No dialog. | blocked: not run (turn budget) |
| FILES-33 | P2 | web | Files survive a reload ([finishes](../files/the-files-workspace.md#finishes)). | Files present. | 1. Reload. | Still listed. | blocked: reload not risked |

Not checkable by hand:

- Whether a connected folder can be disconnected (no UI; [Open questions](../files/the-files-workspace.md#open-questions-and-verification)).
- Whether there should be a write/delete gate for Pi's `files` tool (product call).

## extensions/the-extensions-hub.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| HUB-01 | P2 | web | `/plugins` shows Snake and the install box ([the simple case](../extensions/the-extensions-hub.md#the-simple-case)). | Defaults. | 1. `/plugins`. | "Snake", "../extensions/snake.js · host runtime", "Paste a plugin URL…". | pass |
| HUB-02 | P2 | web | `/tools` opens Connections ([summary](../extensions/the-extensions-hub.md#summary)). | — | 1. `/tools`. | "Connections". | pass |
| HUB-03 | P2 | web | The expanded card lists Commands and Permissions ([opening](../extensions/the-extensions-hub.md#opening)). | — | 1. Click Snake. | `/snake`; one switch per capability. | pass |
| HUB-04 | P1 | web | Disable: "Snake: disabled" and `/snake` leaves the menu ([changes](../extensions/the-extensions-hub.md#changes)). | Snake on. | 1. Switch off.<br>2. Type `/sn`. | Toast; no menu. | pass |
| HUB-05 | P1 | web | Enable: "Snake: enabled" and `/snake` returns ([changes](../extensions/the-extensions-hub.md#changes)). | Snake off. | 1. Switch on.<br>2. Type `/sn`. | Toast; `/snake` in the menu. | pass |
| HUB-06 | P1 | web | Snake's widget shows above the composer by default ([the simple case](../extensions/the-extensions-hub.md#the-simple-case)). | Fresh. | 1. Look above the composer. | Widget. | fail (document was wrong: nothing until `/snake`; corrected) |
| HUB-07 | P2 | web | `/snake` opens the board in the widget slot ([how extensions show up](../extensions/the-extensions-hub.md#how-extensions-show-up)). | — | 1. Run `/snake`. | Board with "R restart · esc quit". | pass |
| HUB-08 | P1 | web | Disabling an extension removes its widgets ([changes](../extensions/the-extensions-hub.md#changes)). | Snake board open. | 1. `/plugins`, switch Snake off.<br>2. Close Settings. | Board gone. | fail (product bug: board stays until Escape) |
| HUB-09 | P2 | web | Escape outside the composer closes Snake's board ([how extensions show up](../extensions/the-extensions-hub.md#how-extensions-show-up)). | Board open. | 1. Focus out of composer; Escape. | Board gone. | pass |
| HUB-10 | P2 | web | Permission toggle toasts ([changes](../extensions/the-extensions-hub.md#changes)). | Card expanded. | 1. Toggle "register model providers" twice. | "Snake: register model providers revoked" / "granted". | pass |
| HUB-11 | P2 | web | Install with an empty box ([changes](../extensions/the-extensions-hub.md#changes)). | — | 1. Install. | "Enter a URL first." | pass |
| HUB-12 | P1 | web | URL install fails on Excel on the web (suspected bug) ([open questions and verification](../extensions/the-extensions-hub.md#open-questions-and-verification)). | — | 1. Paste a URL; Install. | Record what happens. | pass (confirmed: "Install failed: Function window.prompt is not supported."; box cleared) |
| HUB-13 | P2 | web | Uninstall asks; Escape cancels ([changes](../extensions/the-extensions-hub.md#changes)). | Card expanded. | 1. Uninstall.<br>2. Escape. | "Uninstall "Snake"?" / "This removes the plugin and its settings."; Snake kept. | pass |
| HUB-14 | P1 | web | `extensions_manager` installs without approval in Confirm ([how extensions show up](../extensions/the-extensions-hub.md#how-extensions-show-up)). | Confirm. | 1. Ask Pi to install a tiny extension. | No dialog. | blocked: not run (turn budget) |
| HUB-15 | P2 | web | Escape goes back to the root, then closes ([the simple case](../extensions/the-extensions-hub.md#the-simple-case)). | Plugins page. | 1. Escape twice. | Root, then closed. | pass |
| HUB-16 | P1 | desktop | Remote URL install is refused unless the flag is on ([changes](../extensions/the-extensions-hub.md#changes)). | Flag off. | 1. Install an https URL. | Refused. | fail. desktop (macOS): Install with an https URL calls `window.prompt("Extension name:")`, which shows nothing and freezes the taskpane until it is reopened; the remote-URL gate is never reached (`src/commands/builtins/extensions-hub-plugins.ts:296`). Bug. |

Not checkable by hand:

- Permission switches are not enforced unless "Extension permission gates" is on ([Open questions](../extensions/the-extensions-hub.md#open-questions-and-verification)); needs an extension that uses a revoked capability.
- Whether Pi should install extensions without approval (product call).

## extensions/web-search-and-mcp.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| WEB-01 | P2 | web | Connections page sections in order ([the page](../extensions/web-search-and-mcp.md#the-page)). | — | 1. `/tools`. | External tools, Web search, Extension connections, MCP servers, Bridges. | fail (document listed MCP before Extension connections; corrected) |
| WEB-02 | P2 | web | "External tools" is on by default ([the page](../extensions/web-search-and-mcp.md#the-page)). | Defaults. | 1. Read the switch. | On. | pass |
| WEB-03 | P2 | web | Jina card and "No API key" badge ([the page](../extensions/web-search-and-mcp.md#the-page)). | No key. | 1. Read. | "Jina Search (default)", "Fast web search API with free tier.", "No API key". | pass |
| WEB-04 | P1 | web | Web search needs a key even for Jina (suspected bug) ([open questions and verification](../extensions/web-search-and-mcp.md#open-questions-and-verification)). | No key. | 1. Ask Pi to search. | Record what happens. | pass (confirmed: "Web search API key is missing. …") |
| WEB-05 | P2 | web | Scope controls on an unidentified workbook ([the page](../extensions/web-search-and-mcp.md#the-page)). | Signed-out workbook. | 1. Expand the card. | Session toggle on; "Workbook scope unavailable" disabled; Availability "Session". | fail (document said the session toggle starts off; corrected) |
| WEB-06 | P2 | web | Provider change toast ([the page](../extensions/web-search-and-mcp.md#the-page)). | — | 1. Pick Tavily. | "Web search provider set to Tavily". | pass |
| WEB-07 | P1 | web | Save stores the key and the badge turns "Connected" ([the page](../extensions/web-search-and-mcp.md#the-page)). | Fake `jina_` key. | 1. Type; Save. | "Saved Jina API key"; "Connected". | pass |
| WEB-08 | P2 | web | Save with an empty box ([the page](../extensions/web-search-and-mcp.md#the-page)). | No key. | 1. Save. | "Enter an API key first." | pass |
| WEB-09 | P2 | web | The box is prefilled with the masked key ([the page](../extensions/web-search-and-mcp.md#the-page)). | Key saved. | 1. Expand. | `jina******…ef`. | pass |
| WEB-10 | P1 | web | Save without retyping stores the mask (suspected bug) ([open questions and verification](../extensions/web-search-and-mcp.md#open-questions-and-verification)). | WEB-09. | 1. Save.<br>2. Read storage. | Record what happens. | pass (confirmed: mask stored; "⚠️ Jina keys usually start with…") |
| WEB-11 | P2 | web | Validate reports failure ([the page](../extensions/web-search-and-mcp.md#the-page)). | Fake key. | 1. Validate. | "✗ Jina Search (default) search request failed (401): …" | pass |
| WEB-12 | P2 | web | Clear ([the page](../extensions/web-search-and-mcp.md#the-page)). | Key saved. | 1. Clear. | "Cleared Jina API key"; "No API key". | pass |
| WEB-13 | P2 | web | Validate with no key ([the page](../extensions/web-search-and-mcp.md#the-page)). | No key. | 1. Validate. | "No API key to validate." | pass |
| WEB-14 | P2 | web | Master switch toasts ([the page](../extensions/web-search-and-mcp.md#the-page)). | — | 1. Toggle twice. | "External tools disabled" / "enabled". | pass |
| WEB-15 | P2 | web | MCP form errors keep typed values ([adding an mcp server](../extensions/web-search-and-mcp.md#adding-an-mcp-server)). | — | 1. Add with empty name.<br>2. Add with name only. | "Extensions: MCP server name cannot be empty."; values kept (form folds away). | pass |
| WEB-16 | P2 | web | Adding a server ([adding an mcp server](../extensions/web-search-and-mcp.md#adding-an-mcp-server)). | Name + URL. | 1. Add. | "Added MCP server"; card with URL, "(none)", Enabled, Test, Remove. | pass |
| WEB-17 | P1 | web | Pi never gets the `mcp` tool (suspected bug) ([open questions and verification](../extensions/web-search-and-mcp.md#open-questions-and-verification)). | Server added, enabled. | 1. Ask Pi to list its tool names. | Record what happens. | pass (confirmed: no `mcp`) |
| WEB-18 | P2 | web | Test an unreachable server ([adding an mcp server](../extensions/web-search-and-mcp.md#adding-an-mcp-server)). | Server at 127.0.0.1:9. | 1. Test. | "pdtest: Failed to fetch". | pass |
| WEB-19 | P1 | web | Remove deletes without asking ([adding an mcp server](../extensions/web-search-and-mcp.md#adding-an-mcp-server)). | Server card. | 1. Remove. | "Removed MCP server: pdtest"; no dialog. | pass |
| WEB-20 | P1 | web | A failed search shows the tool error and the setup card ([what the user sees when pi uses them](../extensions/web-search-and-mcp.md#what-the-user-sees-when-pi-uses-them)). | No key. | 1. Ask Pi to search once. | Card "Web search "UK base rate""; error; "Web search needs an API key" with Save/Dismiss. | pass |
| WEB-21 | P2 | web | Saving a key in the setup card ([what the user sees when pi uses them](../extensions/web-search-and-mcp.md#what-the-user-sees-when-pi-uses-them)). | Setup card shown. | 1. Enter a fake key; Save. | Ready toast. | fail (document omitted the failure path: key saved, card shows "Key saved. Validation: …"; corrected) |
| WEB-22 | P1 | web | Disclosure-bar switches change nothing (suspected bug) ([the disclosure bar](../extensions/web-search-and-mcp.md#the-disclosure-bar)). | Bar shown. | 1. Customize; switch web off.<br>2. Read settings. | Record what happens. | pass (confirmed: nothing stored) |
| WEB-23 | P2 | web | `web_search` asks no approval in Confirm ([modifiers](../extensions/web-search-and-mcp.md#modifiers)). | Confirm. | 1. Ask Pi to search. | No dialog for the search. | pass |
| WEB-24 | P2 | web | `fetch_page` limits and rate limit ([what the user sees when pi uses them](../extensions/web-search-and-mcp.md#what-the-user-sees-when-pi-uses-them)). | — | 1. Ask Pi to fetch one page twice quickly. | "Rate limited for {host}…" | blocked: not run (turn budget) |
| WEB-25 | P2 | web | Escape goes back, then closes ([the interaction event by event](../extensions/web-search-and-mcp.md#the-interaction-event-by-event)). | Connections. | 1. Escape twice. | Root, then closed. | pass |

Not checkable by hand:

- Which search providers work direct on which host ([Modifiers](../extensions/web-search-and-mcp.md#modifiers)); needs real keys.

## extensions/experimental.md

| ID | P | Device | Claim | Setup | Steps | Expected | Result |
| --- | --- | --- | --- | --- | --- | --- | --- |
| EXP-01 | P2 | web | Page sections, five rows, warnings, "Ready", footer ([the simple case](../extensions/experimental.md#the-simple-case)). | — | 1. `/experimental`. | Word for word as described. | pass |
| EXP-02 | P2 | web | Switch toasts for Dark mode ([the interaction event by event](../extensions/experimental.md#the-interaction-event-by-event)). | — | 1. Toggle twice. | "Dark mode: Enabled" / "Disabled". | pass |
| EXP-03 | P3 | web | Dark mode follows a dark Office theme ([dark mode](../extensions/experimental.md#dark-mode)). | Office dark theme. | 1. Turn on. | Taskpane dark. | blocked: Office theme is light in the orb (taskpane stayed light) |
| EXP-04 | P2 | web | Typed on/off toasts are lower case ([the interaction event by event](../extensions/experimental.md#the-interaction-event-by-event)). | — | 1. `/experimental on dark-mode`, then `off`. | "Dark mode: enabled" / "disabled". | pass |
| EXP-05 | P2 | web | Unknown feature ([the interaction event by event](../extensions/experimental.md#the-interaction-event-by-event)). | — | 1. `/experimental on foo`. | "Unknown feature: foo. Available: …" | pass |
| EXP-06 | P2 | web | `mcp` redirect ([the interaction event by event](../extensions/experimental.md#the-interaction-event-by-event)). | — | 1. `/experimental on mcp`. | "External tools (including MCP) are managed in /tools, not /experimental." | pass |
| EXP-07 | P2 | web | `help` usage ([the interaction event by event](../extensions/experimental.md#the-interaction-event-by-event)). | — | 1. `/experimental help`. | Usage with features list. | pass |
| EXP-08 | P2 | web | Bridge URL `show` ([bridge settings by command](../extensions/experimental.md#bridge-settings-by-command)). | — | 1. `/experimental python-bridge-url show`. | "Python bridge URL: https://localhost:3340". | pass |
| EXP-09 | P3 | web | Bridge token `show` ([bridge settings by command](../extensions/experimental.md#bridge-settings-by-command)). | No token. | 1. `/experimental tmux-bridge-token show`. | Record the message. | pass ("Tmux bridge token is not set. Example: /experimental tmux-bridge-token <token>") |
| EXP-10 | P1 | web | `/experimental` is refused while busy ([summary](../extensions/experimental.md#summary)). | Tab responding. | 1. `/experimental`. | "Can't run /experimental while Pi is busy". | pass |
| EXP-11 | P2 | web | Escape goes back, then closes ([the simple case](../extensions/experimental.md#the-simple-case)). | Page open. | 1. Escape twice. | Root, then closed. | pass |
| EXP-12 | P1 | desktop | Remote-URL gate bypass through the sandbox (suspected bug) ([open questions and verification](../extensions/experimental.md#open-questions-and-verification)). | Flag off. | 1. Install an https extension URL. | Record what happens. | fail. desktop (macOS): same run as HUB-16: the install freezes at the name prompt before any gate, so the bypass cannot be reached on Excel for Mac either. |

Not checkable by hand:

- What the widget v2, permission and sandbox flags change for an installed extension (needs a test extension).

## Results so far

One pass on 2026-10-26 by an agent in an Amp orb, at pi-for-excel `ec453df`. It drove Pi in Excel on the web (Chrome, OfficeOnline, signed-out link workbook) through CDP on the taskpane target: real keystrokes and clicks for typing, Enter and Escape, DOM clicks for buttons, a CDP file chooser for Upload, and synthetic drag events for drops (a real drag through Excel's frame was not possible). Results were read from the DOM, the toasts, Office.js and the add-in's IndexedDB, with a few screenshots. Seven real model turns were used (GPT-6.1 Sol, then GPT-6 Luna). Nothing on Excel desktop was run.

149 items: 118 pass, 9 fail, 22 blocked. Thirteen passes confirm suspected bugs: RULES-06, RULES-09, RULES-24, CONV-06, SKILLS-20, SKILLS-21, FILES-14, FILES-20, HUB-12, WEB-04, WEB-10, WEB-17, WEB-22. CONV-07 settles an open question (Excel on the web accepts the Ratio format). The fails:

- FILES-25: "Connect folder" is offered on Excel on the web but always fails in the cross-origin frame. Product bug.
- HUB-08: disabling an extension leaves its legacy widget (Snake's board) on screen. Product bug.
- FILES-22: "Copy content" fails on the web (clipboard blocked). The document is corrected; it is the same root cause as other clipboard failures on the web.
- FILES-10, FILES-19, HUB-06, WEB-01, WEB-05, WEB-21: the documents were wrong and are corrected.

Blocked items need Excel desktop, a reload of the Excel session (which repeats credential restore), a second taskpane, real search keys or MCP servers, or turns beyond the budget. No document in this cluster is marked verified: P1 and P2 items are still blocked, and the fails are not yet filed in bug-triage.md.
