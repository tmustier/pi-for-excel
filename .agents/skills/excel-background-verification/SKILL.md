---
name: excel-background-verification
description: Run model-driven end-to-end tests in real local Excel while keeping it in the background. Use for local acceptance of behavior changes and dependency upgrades involving workbook tools, taskpane UI, recovery, auth or providers.
---

# Background Excel end-to-end tests

Follow the acceptance policy in `docs/coding-standards.md`. The model performs the work through the taskpane; the test bridge submits prompts and independently reads the workbook.

## Set up

Run from the worktree containing the code under test. You need:

- the sideloaded manifest pointing to `https://localhost:3141/src/taskpane.html`
- trusted `cert.pem` and `key.pem` files in the worktree
- an approved current model with working authentication
- a computer-use tool with Accessibility and Screen Recording permissions

Inspect existing listeners before starting servers. Reuse the correct test setup or stop only processes you own.

```bash
RUN_DIR=$(mktemp -d /tmp/pi-excel-e2e.XXXXXX)
TOKEN=$(node scripts/background-verify-bridge-server.mjs token)

PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost \
  npm run background:verify:bridge > "$RUN_DIR/bridge.log" 2>&1 &
BRIDGE_PID=$!

VITE_PI_BACKGROUND_VERIFY_URL=https://localhost:3157 \
VITE_PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" \
  npm run dev > "$RUN_DIR/vite.log" 2>&1 &
VITE_PID=$!

curl --fail https://localhost:3157/health
```

Wait for server readiness before probing. Keep the bridge loopback-only and tokened; do not print or commit its token. Start `npm run proxy:https` if the configured provider route needs the local CORS proxy (the `openai-codex` route does).

### Open the pane without the foreground

Excel for Mac does not put a sideloaded add-in on the ribbon until it has been activated once in the running process, and the Home-tab "Add-ins" flyout does not render for a background app. Bootstrap through the document instead: a workbook can carry an auto-open task pane part that references the sideloaded manifest (`store="developer" storeType="Registry"`).

```bash
SCRATCH=~/Library/Containers/com.microsoft.Excel/Data/Documents/pi-e2e-scratch.xlsx
python3 .agents/skills/excel-background-verification/scripts/scratch-workbook.py create "$SCRATCH"
open -g -a "Microsoft Excel" "$SCRATCH"      # launch straight into the workbook, in the background
sleep 12; curl --fail https://localhost:3157/health   # expect one client
```

Excel opens the pane on document open and the `Open Pi` ribbon button appears for the rest of the process. The auto-open fires once per Excel process: to reload the pane (for example after switching the Vite worktree) press `Close Pi for Excel` and then `Open Pi` through `computer_use`, passing `element_index` as a string. Keep the scratch workbook inside the Excel container so the sandbox never raises a "Grant File Access" sheet; `scratch-workbook.py tag <existing.xlsx>` adds the part to a workbook Excel has saved.

If `clients` is still empty, inspect Excel with `computer_use` / `sky.get_app_state`. Check the workbook, pane, manifest URL and server before changing caches. For a pane showing a network error, use a semantic accessibility `AXPress` on Try Again after starting the server.

Use background semantic accessibility actions only. Record the foreground app and window before and after the test.

### Before you call it blocked

Everything below has looked like a foreground-only step and was worked around in the background. Report a blocker only after these are exhausted, and say which one you hit.

| Looks like | Is | Do |
|---|---|---|
| No `Open Pi` button on the ribbon after Excel relaunched; manifest is intact | Sideloaded add-ins are absent from the ribbon cache until activated once per Excel process | `scratch-workbook.py create` and `open -g` into it; the pane opens from the document and the button appears |
| Home-tab *Add-ins* flyout does not open (AXPress toggles the button, no popover window) | It does not render for a background app | Not needed; use the document bootstrap above |
| `sky.get_app_state` fails with `cgWindowNotFound`, or Excel lists `AXApplication` as its only window | Seen once right after `open -g` with no document | Wait for the launch to settle and query again; otherwise quit Excel and `open -g` straight into the scratch workbook, which exposed a normal `AXStandardWindow` |
| `computer_use` click does nothing | `element_index` was passed as a number | Pass it as a string, e.g. `"63"` |
| `setTimeout is not defined` inside `computer_use` | The JS sandbox has no timers | Sequence `sky` calls with `await`; wait in the shell (`sleep`) between calls |
| AppleScript `save workbook as` returns `Parameter error (-50)` | Path built from `POSIX path of (path to home folder)` | Pass a literal absolute path string, inside the Excel container |
| Pane cannot be reloaded after switching the Vite worktree | The document auto-open fires once per process | `Close Pi for Excel` then `Open Pi` through `computer_use` |
| Persisted state cannot be checked without trusting the pane | It can | `read-taskpane-setting.py <key>` decodes the record from WebKit's IndexedDB |
| `computer_use` returns `codex app-server exited before returning a response` on every call | The tool's own server is down; Excel is fine | Read the tree with `scripts/axdump.swift` (below); pressing buttons waits for the tool to come back |

Opening a workbook does not activate Excel: both `open -g -a "Microsoft Excel" <file>` and `osascript -e 'tell application "Microsoft Excel" to open POSIX file "…"'` leave the frontmost app unchanged.

## Run the test

Create a test chat through the actual UI. The bridge's `submitPrompt` sends message text, so `/new` and `/name` can reach the model literally rather than dispatching commands.

Define a command shorthand in the same shell:

```bash
bridge() {
  PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost \
    npm run background:verify:command -- "$@"
}

bridge status
bridge selectModel '{"provider":"openai-codex","modelId":"gpt-5.6-sol"}'
```

Use a currently approved model. If health lists multiple clients, pass `--clientId <id>` to target the scratch workbook on each command.

Adapt the prompt to the changed feature. This example exercises model-driven sheet creation, writes, formulas and reads:

```bash
bridge submitPrompt \
  '{"text":"Use workbook tools to create a new sheet _pi_e2e_smoke; stop if it exists. Write 2 and 3 into A1:A2 and =SUM(A1:A2) into A3. Read back the values and formula. Leave the sheet for independent inspection. Do not change other sheets.","waitForIdle":true,"timeoutMs":120000}' \
  --timeout 130000

bridge readRange '{"address":"_pi_e2e_smoke!A1:A3"}'
```

Check actual tool execution and `lastAssistant.model` / `stopReason`. An idle runtime can also mean the request failed. Independently assert that A3 is 5 and retains `=SUM(A1:A2)`; check formatting or objects when relevant.

Clean up through the model, then verify the scratch sheet is gone:

```bash
bridge submitPrompt \
  '{"text":"Delete only _pi_e2e_smoke using your workbook tool. This scratch-sheet deletion is approved. Leave all other sheets unchanged.","waitForIdle":true,"timeoutMs":120000}' \
  --timeout 130000
bridge officeProbe
```

## Diagnose failures

| Symptom | Check |
|---|---|
| No bridge client | Pane not open: bootstrap with `scratch-workbook.py` and `open -g` into the workbook, or press `Open Pi` on the ribbon if the add-in has been activated this process; Vite URL and bridge environment configured |
| Pane does not auto-open from a second tagged workbook | Expected within one Excel process; press `Open Pi` instead, or quit and relaunch Excel into the workbook |
| Pane opens as "New Office Add-in" with an add-in error or endless loading | The webextension reference does not resolve; keep `store="developer" storeType="Registry"` and the manifest `<Id>` |
| Pane loads, model returns `Load failed` | Configured proxy is running and reachable |
| Cannot set the WebKit textarea through accessibility | Submit through the bridge |
| Unsaved workbook has no workbook ID | Identify it through the client and sheet/range evidence; `documentIdentityProbe` shows both host URL sources beside the resolved context |
| Every Office.js call hangs after an AppleScript save | Excel is showing a "Grant File Access" sheet for a folder without a sandbox bookmark (`~/Documents` included); cancel it and save into `~/Library/Containers/com.microsoft.Excel/Data/Documents/` |
| Model turn fails with `Importing a module script failed.` | `npm ci` ran under the live Vite server and wiped `node_modules/.vite/deps`; restart Vite and reload the pane |

Use `officeProbe`, `readRange`, `readUsedRange` and `listCharts` for independent inspection. `workbookWriteProbe` is a reversible Office.js diagnostic, not a model test. `writeRange` and `clearRange` can prepare fixtures, but must not perform the work the model is meant to do. `configureProxy` changes saved settings; restore them after transport-specific tests.

For anything the taskpane persists (recovery snapshots, files-workspace metadata and audit trail, settings), read the record straight out of Excel's WebKit IndexedDB rather than through the pane:

```bash
python3 .agents/skills/excel-background-verification/scripts/read-taskpane-setting.py --list
python3 .agents/skills/excel-background-verification/scripts/read-taskpane-setting.py workbook.recovery-snapshots.v1
```

It copies the database and its WAL and decodes one record to JSON, so a before/after diff shows exactly which entries a change dropped, kept or rewrote.

Rendered tool cards are visible in the pane's accessibility tree (`HTML content` under the `Pi for Excel` container), which is how to assert diff tables, trees and images. The dump truncates on long chats; scroll the `HTML content` element with `sky.scroll` or start a fresh chat with `submitInput` `{"text":"/new"}`. Slash commands such as `/new` and `/history` go through the real composer with `submitInput`; `submitPrompt` sends text to the model.

When `computer_use` is unavailable, the same tree is readable without it. System Events stops at the web area, so use the `AXUIElement` walker:

```bash
swiftc -O -o /tmp/axdump .agents/skills/excel-background-verification/scripts/axdump.swift
/tmp/axdump "Microsoft Excel" > "$RUN_DIR/pane-ax.txt"
rg -n 'AXTable|AXLink|AXPreformattedStyleGroup|value="' "$RUN_DIR/pane-ax.txt" | tail -60
```

It prints role, subrole, title, value, description and URL per element, indented by depth, and does not press anything.

## Report and clean up

Record the tested revision, provider/model, thinking level, prompt, observed tools, independent assertions, cleanup and unchanged foreground app/window. State untested provider or host paths. Keep useful logs and screenshots without credentials.

Stop the test servers and any proxy you started. If the user needs a running development environment, explicitly hand over the normal dev server and proxy; stop the test bridge and remove its token.

```bash
kill "$BRIDGE_PID" "$VITE_PID"
unset TOKEN
```

Confirm the owned listeners have exited, then remove the worktree when it is no longer serving the user.
