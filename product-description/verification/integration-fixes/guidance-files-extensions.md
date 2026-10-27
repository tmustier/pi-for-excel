# Integration verification: guidance, files and extensions

This file reruns every item in [../guidance-files-extensions.md](../guidance-files-extensions.md) against the combined fix build, branch `integration-fixes`. That branch is `main` plus PRs #786–#799 (see `INTEGRATION.md` on the branch). The session was set up at `a2b3f2d`. The branch was then force-pushed several times, so the full pass was run on the final stacked head, `2d91e58` (top of #800, ux-system/#793), with the dev server restarted from that commit. Every row below is the result at `2d91e58`. The run was on 27 October 2026, by an agent in an Amp orb.

Expected behavior is the checklist as amended by the fix PRs and `goal.md` "Fixes in flight". Where a PR deliberately changed behavior, the PR wins, and the note names it ("pass (#NNN changed)"). For this cluster the deciding rule is #790's "Consent outside the workbook" in `src/tools/DECISIONS.md`:

- Confirm mode asks before every `files` write or delete and every changing `extensions_manager` action.
- Auto mode still asks before new code runs and before irreversible loss.
- `instructions`, `conventions` and `skills` never ask.

How it ran:

- **web**: Pi in Excel on the web through `node scripts/orb/excel-web-session.mjs --check` (host Excel, platform OfficeOnline, check ok), signed-out link workbook (no workbook identity).
  - The taskpane (349×622) was driven over CDP on its own target: `Input.insertText` and key events for typing, Enter and Escape; DOM clicks for buttons; `DOM.setFileInputFiles` for Upload; synthetic `DragEvent`s for drops.
  - Results were read from the DOM, a toast observer, Office.js and the add-in's IndexedDB.
  - Model requests were captured by wrapping `fetch` in the taskpane. The live tool set was read from the newest `additional_tools` entry in each request.
- **Second taskpane**: the same workbook link opened in a second Chrome tab of the same profile, with Pi opened from the ribbon.
- **Reload**: `location.reload()` in the taskpane target.
- **Mocks and local servers**:
  - s.jina.ai, api.tavily.com and a test site (`pd-mock.example`) were mocked with CDP `Fetch` on the taskpane target.
  - MCP: `@modelcontextprotocol/server-everything@2.0.1 streamableHttp` on `localhost:3901`.
  - Remote extension: a local server on `localhost:3902` that logged every request.
- **Clipboard refusal** (FILES-22): `navigator.clipboard.writeText` was made to reject and `document.execCommand` to return false.
- **Model turns** ran on GPT-6 Luna (GPT-6.1 Luna is not in the registry). The shared `PI_EXCEL_WEB_AUTH_JSON` login was never logged out or refreshed.
- **Cleanup**: every test file, rule, convention, skill, key, MCP server, extension and session integration added was removed afterwards.

**Result at `2d91e58`: 156 items. 152 pass, 1 fail, 3 blocked.**

- The fail is WEB-31, a document error: the setup card closes itself after a good key. That code is unchanged since `main`.
- Blocked: RULES-25 and FILES-24 are desktop only (skipped). EXP-03 was not rerun, because it needs the standalone taskpane with dark media.

What changed against `main`:

- **Every item that failed on `main` now passes**, or behaves as its owning PR describes: RULES-26, RULES-27 (#788), SKILLS-22, HUB-08, HUB-16 on the web, FILES-25 (#790) and FILES-22 (#795 `copyText`). The document-error rows stay as corrected.
- **No item that passed on `main` regressed.**
- **Four new bugs**, listed below. Three are #790's and one is a #790 copy error; none is a checklist regression.
  - The consent dialog has no "Stop Pi" (combination).
  - Rename and Delete show a false "That file is no longer available." toast.
  - Turning on MCP for a session turns web search off.
  - The web-search SKILL.md contradicts itself about Jina keys.
  - #790's owner confirmed the first three. The integration thread holds local fixes for the first two (717595db, then the toast fix) and has the third's patch.

The confirmed suspected bugs from `main` are now fixed as their PRs describe:

- #788: RULES-09 and RULES-24, CONV-06 and FILES-14.
- #790: RULES-06, SKILLS-20/21, FILES-20, HUB-12, HUB-14, WEB-04, WEB-10, WEB-17 and EXP-12.
- #793: WEB-22.

## Results

| ID | Result | Note |
| --- | --- | --- |
| RULES-01 | pass | `/rules`: Rules page, "All my files" selected (`aria-selected`), focus in the textarea. |
| RULES-02 | pass | Hint, placeholder "Your preferences and habits, e.g. …", "0 / 2,000 chars". |
| RULES-03 | pass | Gear → "Rules" opens the page with focus in the textarea. |
| RULES-04 | pass | Settings row "Rules & conventions" opens it. The row now reads "On: All my files" (#790). |
| RULES-05 | pass (#788 changed) | "Workbook: Current workbook" and "Can't identify this workbook right now, so rules for this file can't be saved. Try saving the file first." The box is `disabled`; typing adds nothing; "0 / 4,000 chars". |
| RULES-06 | pass (#790 changed) | Save: toast "Rules saved", and the overlay closes (does not go back to Settings). Saving with nothing changed says "No changes to save". |
| RULES-07 | pass (#790 changed) | Cancel with no changes closes the overlay silently. |
| RULES-08 | pass | The first Escape in the box only leaves it; the page stays open. |
| RULES-09 | pass (#788 changed) | Workbook rules can no longer be typed on an unidentified workbook (see RULES-05), so nothing is dropped silently. |
| RULES-10 | pass (wording) | "Discard changes?", "You have unsaved changes to your rules or formats. Discard them?", "Keep editing", "Discard". |
| RULES-11 | pass | Escape on the question closes it; the page stays with the draft (11 chars). |
| RULES-12 | pass | × with changes asks; "Discard" closes the overlay; nothing stored. |
| RULES-13 | pass | A Number dp change on Formats, then "All my files" and Escape: "Discard changes?". |
| RULES-14 | pass | "2,050 / 2,000 chars", class `is-warning` (red); Save stored all 2,050 characters. |
| RULES-15 | pass | Spaces saved: stored `user.instructions` removed; box empty on reopen. |
| RULES-16 | pass | Card "Remember user rules", result "Updated user rules (23/2000 chars):". |
| RULES-17 | pass | No toast during or after Pi's `instructions` call. |
| RULES-18 | pass | Confirm mode: `instructions` append ran with no dialog. |
| RULES-19 | pass | Card result "Error: workbook identity unavailable. Can't update workbook rules right now." |
| RULES-20 | pass | IndexedDB `settings` → `user.instructions` held the saved text. |
| RULES-21 | pass | `/rules` opened about 4 s into a long turn; the turn kept running. |
| RULES-22 | pass (#793 changed) | One Escape on the page (opened with `/rules`) closes it; the turn kept running. A second Escape after that reaches the composer and stops the turn ("Request aborted"). That is the documented rule (`turn/stopping.md` "A second Escape after a claiming surface has closed stops the turn", and the edge case "Pressing Escape twice quickly"), so the item's "Escape twice" step should become "Escape once". |
| RULES-23 | pass | " draftZZ" typed, `location.reload()`: the draft is gone, with no prompt. |
| RULES-24 | pass (#788 changed) | Page open on "Pane rule KIWI"; Pi (pane B) appended "Prefer GBP."; Save unchanged → "No changes to save", and storage kept both lines. With the same box edited on the page, Save asks "Rules changed" / "Pi changed these rules while this page was open. Save anyway and replace that change with yours?": "Keep editing" kept the draft and wrote nothing; "Replace" stored "Mine only". |
| RULES-25 | blocked (desktop) | Desktop only; not run. |
| RULES-26 | pass (web) | Two taskpanes on the same link workbook: "Pane rule KIWI" saved in pane A appeared in pane B's next model request (newest rule after the older "ZQY"), without focusing pane A. The desktop part is not rerun. |
| RULES-27 | pass (#788) | The saved rule was in each next request: "Cross pane rule ZQX" right after Save; again after a reload; "Fresh rule ZQY" after a later Save. |
| CONV-01 | pass | Hint and six cards with the default strings. |
| CONV-02 | pass | Format field, three previews, dp/neg/zero selects and the symbol field. |
| CONV-03 | pass | dp 0 and € gave `€* #,##0_);€* (#,##0);€* --_)` and previews "€1,235", "€(1,235)", "€--". |
| CONV-04 | pass | With Currency saved at dp 1, Pi's "currency preset" set A1:A3 to exactly `€* #,##0.0_);€* (#,##0.0);€* --_)`; Excel shows "€1,234.5", "€(1,234.5)", "€--". |
| CONV-05 | pass | Default previews "$1,234.50", "$(1,234.50)", "$--". |
| CONV-06 | pass (#788 changed) | Percent previews "12.5%", "(12.5%)", "--%" (sample 0.125, ×100). |
| CONV-07 | pass | Ratio format on 1234.5, −1234.5, 0: "1,234.5x", "(1,234.5x)", "--". |
| CONV-08 | pass | Typed `0.000`: quick options gone, previews "Custom", reset link shown. |
| CONV-09 | pass | The reset link restored `#,##0.00_);(#,##0.00);--_)`. |
| CONV-10 | pass | "custom-1" with Name, Description, Format, "Custom" previews, "Remove preset". |
| CONV-11 | pass | `notacolor` in Hardcoded values: the field shows `#0000ff` again; no toast. |
| CONV-12 | pass | Cancel → "Discard changes?" → Discard: no custom preset in `conventions.v1`. |
| CONV-13 | pass | After Currency dp changed to 2 and saved, A1:A3 kept `€* #,##0.0_)…`. |
| CONV-14 | pass | Confirm mode: no dialog; card "Update conventions". |
| CONV-15 | pass | Renaming custom-2 to "custom-1" was ignored; no message. |
| SKILLS-01 | pass | "5 skills active (5 bundled, 0 external)". |
| SKILLS-02 | pass | Five bundled skills with "Bundled" badges; the web-search description now says every provider needs a key (#790). |
| SKILLS-03 | pass | "No external skills installed." / "Paste a SKILL.md below to add one." |
| SKILLS-04 | pass | "Paste a SKILL.md document first." |
| SKILLS-05 | pass | "Extensions: Invalid SKILL.md document: expected frontmatter with name and description." |
| SKILLS-06 | pass | "Extensions: Skill name cannot contain path separators." |
| SKILLS-07 | pass | "Saved external skill: pd-test"; box empty; "6 skills active (5 bundled, 1 external)". |
| SKILLS-08 | pass | "Disabled skill: web-search"; "5 skills active (4 bundled, 1 external)". |
| SKILLS-09 | pass | "Enabled skill: web-search". |
| SKILLS-10 | pass | Location `skills/external/pd-test/SKILL.md`, "Remove". |
| SKILLS-11 | pass (#788 changed) | Remove now asks "Remove skill "pd-test"?" / "This deletes its SKILL.md from Files. Pi can no longer use it."; Escape kept it; Remove → "Removed external skill: pd-test". |
| SKILLS-12 | pass | Reinstalled pd-test came back switched off and not counted. |
| SKILLS-13 | pass | `name: web-search`: "Shadowed" badge, switch disabled. |
| SKILLS-14 | pass | Files: SKILLS 2 → pd-test → SKILL.md, web-search → SKILL.md. |
| SKILLS-15 | pass (#793 changed) | Opened with `/skills`: the first Escape closes it, because the shell closes on the page it opened at. |
| SKILLS-16 | pass | Card "List skills". |
| SKILLS-17 | pass | "Read skill web-search (bundled)", then "Read skill web-search (bundled) (cached)". |
| SKILLS-18 | pass | "Skill not found: `pd-test`. Available: …" for the switched-off skill. |
| SKILLS-19 | pass | Auto: "Install skill pd-pi", no dialog (`skills` is not gated under #790). |
| SKILLS-20 | pass (#790/#794 changed) | Confirm mode: `skills` list asked nothing. |
| SKILLS-21 | pass (#790/#794 changed) | Confirm mode: `skills` read asked nothing. |
| SKILLS-22 | pass (#790) | Pi wrote `skills/pd2/SKILL.md`; the Skills page lists pd2 as External ("7 skills active (5 bundled, 2 external)"). |
| FILES-01 | pass | Files / "Documents available to Pi" with Upload and Connect folder. |
| FILES-02 | pass | Gear → Files opens it. |
| FILES-03 | pass (#793 changed) | The overlay now takes focus (the filter box). Escape, Escape: closed, focus back in the composer. |
| FILES-04 | pass | `/files` with the overlay open closed it. |
| FILES-05 | pass | Synthetic drag: `is-drag-over`, "Drop files to import to file workspace". |
| FILES-06 | pass | "Imported 1 file into Files."; draft "draft text" unchanged; nothing sent. |
| FILES-07 | pass | "Imported 2 files into Files." |
| FILES-08 | pass | Upload (`multiple`), two files: "Imported 2 files." |
| FILES-09 | pass | "{size} · Uploaded · just now", newest first. |
| FILES-10 | pass | No "Workbook" badge on the unidentified workbook (document as corrected). |
| FILES-11 | pass | BUILT-IN DOCS 47, "Read only" / "Pi documentation". |
| FILES-12 | pass | SKILLS section with folder rows (see SKILLS-14). |
| FILES-13 | pass | "53 files · 363 KB · Sandboxed workspace" with 6 user files: built-in docs are counted. |
| FILES-14 | pass (#788 changed) | The second `budget.csv` was kept as `budget (2).csv`: "Imported 1 file into Files. budget.csv already existed, so the new one was saved as budget (2).csv." |
| FILES-15 | pass | `BUD` → budget (2).csv, budget.csv. |
| FILES-16 | pass | "No matching files" / "Try a different filter term." |
| FILES-17 | pass | The first Escape in the filter left the box; the overlay stayed. |
| FILES-18 | pass | "Text · 26 B · Uploaded · just now", heading rendered, Open ↗, Download, Rename, Delete. |
| FILES-19 | pass | compaction.md preview without "(preview truncated)" (document as corrected). |
| FILES-20 | pass (#790 changed) | Escape in a preview returns to the list; the second Escape closes. |
| FILES-21 | pass | "Built-in documentation — read only. Pi references this automatically."; Copy content, Download. |
| FILES-22 | pass (#795 changed; P-166) | Copy content: "Copied anti-slop-policy.md to clipboard.". With the Clipboard API rejecting and `execCommand` returning false (each called once), `copyText` opened "Copy text" / "This window can't write to the clipboard. The text is selected: press Ctrl+C (⌘C on Mac).": a read-only textarea, focused, 0–4323 of 4323 characters selected, "Done"; Escape returned to the preview. |
| FILES-23 | pass | Open ↗ opened a `blob:https://localhost:3141/…` window. |
| FILES-24 | blocked (desktop) | Desktop only (Excel for Mac); not run. |
| FILES-25 | pass (#790 changed) | Connect folder is `disabled`, with "Connecting a local folder isn't possible in Excel on the web: …" |
| FILES-26 | pass | "Rename file", "notes.md — leave off the extension to keep it.", prefilled, focused. |
| FILES-27 | pass, new bug | "Renamed to docs/renamed.md."; the preview title is renamed.md. A second toast, "That file is no longer available.", also appears (bug 2). |
| FILES-28 | pass, new bug | "Delete file?" with `docs/renamed.md`; "Deleted renamed.md."; list. The same false second toast appears (bug 2). |
| FILES-29 | pass | ← returns to the list. |
| FILES-30 | pass | A drop while Pi waited on a consent dialog mid-turn: "Imported 1 file into Files."; the turn continued. |
| FILES-31 | pass | A drop on the transcript (`.pi-messages`) imported nothing. |
| FILES-32 | pass (#790 changed) | Confirm: Pi's write asked "Create a file?" (Create) and its delete asked "Delete a file?" (Delete); both ran after approval. Auto: Pi wrote and rewrote its own `scratch/auto.txt` with no question; deleting the user's `up-one.txt` asked, and Decline left it ("Files delete of 'up-one.txt' cancelled by user."). |
| FILES-33 | pass | Files still listed after `location.reload()`. |
| HUB-01 | pass | Snake, "../extensions/snake.js · host runtime", "Paste a plugin URL…". |
| HUB-02 | pass | `/tools` opens Connections. |
| HUB-03 | pass | `/snake` and one switch per capability. |
| HUB-04 | pass | "Snake: disabled"; no `/snake` in the menu for `/sn`. |
| HUB-05 | pass | "Snake: enabled"; `/snake` back in the menu. |
| HUB-06 | pass | Widget slot empty until `/snake` (document as corrected). |
| HUB-07 | pass | `/snake` opened the board ("Score: 0", "arrows · esc quit" during play). |
| HUB-08 | pass (#790) | Snake switched off with the board open: after Settings closed, the widget slot was empty. |
| HUB-09 | pass | Escape out of the composer closed the board; no turn affected. |
| HUB-10 | pass | "Snake: register model providers revoked" / "granted". |
| HUB-11 | pass | "Enter a URL first." |
| HUB-12 | pass (#790 changed) | URL install on the web works: an in-pane "Install plugin from URL" dialog (source, default permissions, name field) → "Installed: pd-url", sandbox iframe. With the flag off, a toast comes first (HUB-16). |
| HUB-13 | pass | "Uninstall "pd-url"?" / "This removes the plugin and its settings."; Escape kept it; Uninstall → "Uninstalled: pd-url". |
| HUB-14 | pass (#790 changed) | `extensions_manager install_code` now asks, even in Auto: "Install an extension?" with code and "Extra permissions requested: none". Install → enabled with the 8 default capabilities only. Card "Install extension pd-ok-ext". Uninstall asked "Uninstall an extension?" in Auto. |
| HUB-15 | pass (#793 changed) | Opened with `/plugins`: the first Escape closes. Opened from Settings (gear → Extensions): back, then closed. |
| HUB-16 | pass (web, #790) | Flag off: "Remote extension URLs are off. Turn on "Remote extension URLs" in /experimental, then install again."; the local server logged 0 requests. The desktop part is not rerun. |
| WEB-01 | pass | External tools, Web search, Extension connections, MCP servers, Bridges. |
| WEB-02 | pass | External tools on. |
| WEB-03 | pass (#790 wording) | "Jina Search (default)", "Fast web search API. Needs a free API key from jina.ai.", "No API key". |
| WEB-04 | pass (#790 changed) | A key is required and the copy now says so: "Error: Web search API key is missing. Open /tools and set the Jina API key." |
| WEB-05 | pass | Scope controls (collapsed disclosure): "Enable for this session" on, "Workbook scope unavailable" disabled; Availability "Session". |
| WEB-06 | pass | "Web search provider set to Tavily". |
| WEB-07 | pass | "Saved Jina API key"; "Connected". |
| WEB-08 | pass | "Enter an API key first." |
| WEB-09 | pass (#790 changed) | The field is empty, with placeholder "Saved key jina**********************ef — type a new key to replace it". |
| WEB-10 | pass (#790 changed) | Save without retyping: "Enter an API key first."; `connections.store.v1` kept the real key. |
| WEB-11 | pass (#790) | Validate with the field empty tested the saved key (Bearer `jina_fake…` sent); mocked 401: "✗ Jina Search (default) search request failed (401): …". |
| WEB-12 | pass | "Cleared Jina API key" (also "Cleared Tavily API key"); "No API key". |
| WEB-13 | pass | "No API key to validate." |
| WEB-14 | pass | "External tools disabled" / "External tools enabled". |
| WEB-15 | pass | "Extensions: MCP server name cannot be empty."; the name was kept while the form folded away. |
| WEB-16 | pass | "Added MCP server"; card with URL, Token "(none)", Enabled, Test, Remove. |
| WEB-17 | pass (#790) | The new "Let Pi use MCP servers" session switch ("MCP servers enabled for this session") gave Pi `mcp` with no reload. But it turned web search off (bug 3). |
| WEB-18 | pass | "pdtest: Failed to fetch". |
| WEB-19 | pass | "Removed MCP server: pdtest"; no dialog. |
| WEB-20 | pass | Card "Web search "UK base rate latest Bank of England"", the error, and "Web search needs an API key" with Save/Dismiss. |
| WEB-21 | pass | The mock refused the key: "Key saved. Validation: Jina Search (default) search request failed (401): …" in the card. |
| WEB-22 | pass (#793 changed) | The disclosure bar no longer has dead switches: "Got it" / "Review in Settings" (opens Settings). |
| WEB-23 | pass | Confirm: `web_search` and `fetch_page` asked nothing. |
| WEB-24 | pass | Mock host, parallel pair: "Rate limited for pd-mock.example. Please wait 1000ms between requests." |
| WEB-25 | pass (#793 changed) | `/tools`: the first Escape closes. From Settings: back, then closed. |
| WEB-26 | pass | "everything: reachable (16 tools, direct)". |
| WEB-27 | pass (#790 changed) | No storage edit needed: the session switch added `mcp`, and Pi's "MCP call echo" returned "Echo: pd-ping". |
| WEB-28 | pass | Server switched off: toast "everything: disabled"; Pi's call: "Error: MCP server "everything" is disabled.". Test still probes the disabled server ("reachable (16 tools)"), in the same code as `main` (`extensions-hub-connections.ts:651-664`); the "everything: disabled" in the `main` result is the switch toast. |
| WEB-29 | pass (mocked) | "✓ Jina Search (default) key is valid (1 result, direct)."; "Web search via Jina Search (default)" with [1] [2]; "Fetched page content". |
| WEB-30 | pass (mocked) | Tavily 500 → "… Check your Tavily API key in /tools." warning and "Web search via Jina Search (default)". |
| WEB-31 | fail (document) | Good key in the setup card: "✓ Web search is ready — ask the assistant to try again.", and the card closes itself at once (`markDone` → `onDismiss`, `src/ui/web-search-setup-card.ts:228-231`, unchanged since `main`). "Card stays until Dismiss" is wrong; correct the document. |
| EXP-01 | pass (#790 changed) | Four rows now (Dark mode, Extension widget API v2, Remote extension URLs, Extension sandbox rollback): the extension-permissions flag is removed. Warnings, "Ready", footer present. |
| EXP-02 | pass (#790 wording) | "Dark mode: enabled" / "Dark mode: disabled" (one key, lower case, as typed). |
| EXP-03 | blocked | Not rerun: needs the standalone taskpane with dark media (Excel on the web reports a white theme). |
| EXP-04 | pass | "Dark mode: enabled" / "disabled". |
| EXP-05 | pass | "Unknown feature: foo. Available: dark-mode, remote-extension-urls, extension-sandbox-rollback, extension-widget-v2". |
| EXP-06 | pass | "External tools (including MCP) are managed in /tools, not /experimental." |
| EXP-07 | pass | Usage with the features list. |
| EXP-08 | pass | "Python bridge URL: https://localhost:3340". |
| EXP-09 | pass | "Tmux bridge token is not set. Example: /experimental tmux-bridge-token <token>". |
| EXP-10 | pass (#789 wording) | "Wait for this tab to finish before running /experimental"; the draft is kept. |
| EXP-11 | pass (#793 changed) | `/experimental`: the first Escape closes. |
| EXP-12 | pass (web, #790 changed) | The bypass is closed: with the flag off, install is refused (0 requests). An entry installed with the flag on stays unloaded after the flag is turned off and the pane reloads (no `/pdurl`, 0 requests). The desktop part is not rerun. |

## Regressions and new bugs

No checklist item that passed on `main` regressed. New bugs:

1. **The files/extensions consent dialog has no "Stop Pi", and Stop is inert behind it** (combination; #790 + #794).
   - Repro: Confirm mode; ask "Use the files tool to write a new file scratch/pd-stop.txt containing the word hello." The dialog "Create a file?" has only ×, Decline and Create, and the taskpane's Stop button is `inert`. The same happens for `extensions_manager install_code` ("Install an extension?").
   - Withdrawal itself works: with `inert` removed and Stop clicked, the dialog closed, the card read "Aborted / Request aborted", and no file or extension was created.
   - Cause: `src/taskpane/init.ts:1155-1161` passes `signal` but not `onStop: stopRuntimeTurn` (the Office.js and Python approvals at `:1137`/`:1151` pass it). Introduced in the stack resolution 9ac8d2c.
   - Owner #790; held fix 717595db in the integration thread.
2. **Rename or Delete from a Files preview also toasts "That file is no longer available."** (#790, low).
   - Repro: drop notes.md; open it; Rename to `docs/renamed` → "Renamed to docs/renamed.md." plus "That file is no longer available.". Delete → "Deleted renamed.md." plus the same toast.
   - Cause: `renameFile`/`deleteFile` dispatch the workspace change event before the dialog's after-rename/delete update, so `onWorkspaceChanged` (`src/ui/files-dialog.ts:1001-1022`) still holds the old `detailFileRef` and toasts at `:1016`.
   - Probably also on `main`, and visible now that #793 stacks toasts.
   - Owner #790 (confirmed); fix held in the integration thread.
3. **Turning on "Let Pi use MCP servers" for the session turns web search off** (#790, medium).
   - Repro: identity-less workbook, fresh session (web search on by default). /tools → Let Pi use MCP servers → Scope controls → Enable for this session.
   - Result: storage holds `integrations.session.v1.<id> = ["mcp_tools"]`. The next request carries "The following tools are no longer available: web_search, fetch_page", Pi says web search is unavailable, and the web search session switch reads off.
   - Cause: `setIntegrationEnabledInScope` (`src/integrations/store.ts:175-177`) seeds an unconfigured scope with the defaults only for `workbook`. The read path (`resolveConfiguredIntegrationIds`) applies them to a session when the workbook has no identity. The code dates from `main`; #790's switch exposes it.
   - Owner #790 (confirmed); patch with a regression test sent to the integration thread.
4. **The bundled web-search SKILL.md contradicts itself** (#790 copy, low).
   - The description (line 3) and provider table (line 25) say every provider, Jina included, needs a key. Setup step 4 (`skills/web-search/SKILL.md:44`) still says "Jina works by default with no configuration." Pi read this in the SKILLS-17 turn.
   - Owner #790 (11d3b9d4 edited the file).

## Desktop items for the Mac

- RULES-25 (first save with the Rules page open).
- RULES-26, desktop part (two windows).
- FILES-24 (Download / Open ↗ in a preview; #790's Mac `data:` path was dropped for #796's blob downloads, so recheck).
- HUB-16 and EXP-12, desktop part (the URL install now uses an in-pane name dialog instead of `window.prompt`; check it no longer freezes and the remote-URL gate refuses with the flag off).
