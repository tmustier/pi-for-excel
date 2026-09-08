---
name: excel-background-verification
description: Run model-driven end-to-end tests in real local Excel while keeping it in the background. Use for local acceptance of behavior changes and dependency upgrades involving workbook tools, taskpane UI, recovery, auth or providers.
---

# Background Excel end-to-end tests

Follow the acceptance policy in `docs/coding-standards.md`. The model performs the work through the taskpane; the test bridge submits prompts and independently reads the workbook.

## Set up

Run from the worktree containing the code under test. You need:

- a scratch workbook open in Excel with the Pi taskpane loaded
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

Wait for server readiness before probing. Keep the bridge loopback-only and tokened; do not print or commit its token. Start `npm run proxy:https` if the configured provider route needs the local CORS proxy.

If `clients` is empty, inspect Excel with `computer_use` / `sky.get_app_state`. Check the workbook, pane, manifest URL and server before changing caches. For a pane showing a network error, use a semantic accessibility `AXPress` on Try Again after starting the server.

Use background semantic accessibility actions only. If an action requires foreground focus or raw keyboard/mouse input, report the blocker. Record the foreground app and window before and after the test.

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
| No bridge client | Workbook and Pi pane open; Vite URL and bridge environment configured |
| Pane loads, model returns `Load failed` | Configured proxy is running and reachable |
| Cannot set the WebKit textarea through accessibility | Submit through the bridge |
| Unsaved workbook has no workbook ID | Identify it through the client and sheet/range evidence |

Use `officeProbe`, `readRange`, `readUsedRange` and `listCharts` for independent inspection. `workbookWriteProbe` is a reversible Office.js diagnostic, not a model test. `writeRange` and `clearRange` can prepare fixtures, but must not perform the work the model is meant to do. `configureProxy` changes saved settings; restore them after transport-specific tests.

## Report and clean up

Record the tested revision, provider/model, thinking level, prompt, observed tools, independent assertions, cleanup and unchanged foreground app/window. State untested provider or host paths. Keep useful logs and screenshots without credentials.

Stop the test servers and any proxy you started. If the user needs a running development environment, explicitly hand over the normal dev server and proxy; stop the test bridge and remove its token.

```bash
kill "$BRIDGE_PID" "$VITE_PID"
unset TOKEN
```

Confirm the owned listeners have exited, then remove the worktree when it is no longer serving the user.
