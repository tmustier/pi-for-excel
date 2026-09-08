---
name: excel-background-verification
description: Run real taskpane prompt to model to workbook-tool end-to-end acceptance tests in local Excel without taking foreground focus. Use for local verification of behavior changes and dependency upgrades, including tools, recovery, UI, auth, providers and Office.js paths. Direct Office.js probes and browser fallback checks are supporting evidence only.
---

# Excel Background Verification

Use this skill to verify work in **real Excel with the real pi-for-excel taskpane** without taking over the user's foreground app.

The workflow has two channels:

1. **Computer-use observation** verifies the Excel window/taskpane visually and semantically while asserting the foreground app does not change. Use tools actually exposed in the current session, such as `computer_use` with `sky.get_app_state`.
2. **Taskpane background bridge** submits real prompts through the sidebar and independently reads results from inside Excel via a tokened loopback HTTPS server. Direct Office.js probes diagnose setup; they do not replace model-driven testing.

## End-to-end acceptance gate

For behavior changes and dependency upgrades, require all of the following before claiming local acceptance:

1. Record the tested source revision or tree. Select a current approved provider/model through the actual taskpane and record its ID and thinking level.
2. Submit a real prompt through the sidebar. Let the model execute the workbook tools; do not perform the requested edits through test helpers.
3. Observe actual tool execution and a completed model response. Bridge `ok: true` and `idle: true` alone do not prove success; inspect the response stop reason and errors.
4. Independently read the workbook and assert expected values, formulas and relevant formatting or objects. Do not rely on the model's self-report.
5. Clean up isolated scratch data and independently confirm cleanup. Verify the foreground app and window stayed unchanged.
6. Report the provider/model, thinking level, prompt, tools observed, assertions and remaining coverage gaps. One Codex test does not establish Anthropic or WPS coverage.

Keep unit tests, builds, CI, screenshots and direct host probes as supporting checks. If the real path cannot run, report the blocker and hold acceptance unless the user explicitly waives it. Documentation-only changes need no runtime test.

Do not use raw keyboard/mouse/coordinate actions in this workflow. If a step requires foreground focus or raw input, stop and report that the background lane cannot verify that path.

## Prerequisites

- Excel has the dev manifest sideloaded, pointing to `https://localhost:3141/src/taskpane.html`.
- `cert.pem` and `key.pem` exist in the repo root and are trusted by the local WebView.
- The available computer-use tool has macOS Accessibility and Screen Recording permissions. Do not assume a nested Pi session exposes older `list_apps` / `observe` tool names.
- Run from the repo root.
- Start the normal HTTPS CORS proxy when the configured provider route needs it (`npm run proxy:https`). A loaded pane with a stopped proxy can return `Load failed` on real model requests.

## Start the bridge + dev server

```bash
TOKEN_FILE=$(mktemp /tmp/pi-background-verify-token.XXXXXX)
chmod 600 "$TOKEN_FILE"
TOKEN=$(node scripts/background-verify-bridge-server.mjs token)
printf '%s' "$TOKEN" > "$TOKEN_FILE"

PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" \
PI_BACKGROUND_VERIFY_HOST=localhost \
  npm run background:verify:bridge \
  > /tmp/pi-background-verify-bridge-server.log 2>&1 &
echo $! > /tmp/pi-background-verify-bridge-server.pid

VITE_PI_BACKGROUND_VERIFY_URL=https://localhost:3157 \
VITE_PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" \
  npm run dev \
  > /tmp/pi-for-excel-vite-background-verify.log 2>&1 &
echo $! > /tmp/pi-for-excel-vite-background-verify.pid
```

Check server health:

```bash
curl -k https://localhost:3157/health
```

If `clients` is empty, first inspect the real Excel window. A workbook and loaded Pi pane are required; the startup window alone is insufficient. Confirm the sideloaded URL and running HTTPS server before changing manifests or caches. If the pane shows a network error, use a background semantic `AXPress` on **Try Again** after starting the server. Otherwise close/reopen the pane with semantic controls. Never use raw input fallbacks or focus Excel.

If `clients` lists more than one taskpane, pass `--clientId <id-from-health>` to mutating `background:verify:command` calls so writes target the intended workbook. The server refuses untargeted commands when multiple live clients are connected.

## Computer-use observe lane

Prefer the current session's signed `computer_use` observation surface. Observe with `sky.get_app_state`, and use semantic accessibility actions only when background behavior is guaranteed. Do not assume a general click operation is focus-safe.

For an older installation that actually exposes the following tools, this observation-only invocation is an alternative:

```bash
PI_COMPUTER_USE_STEALTH=1 PI_COMPUTER_USE_STRICT_AX=1 \
pi --tools list_apps,list_windows,observe,search_ui,inspect_ui \
  --no-context-files --no-skills \
  -p 'Use only computer-use tools. Do not focus any window and do not call act. Report the frontmost app, observe Microsoft Excel, and find evidence of Pi for Excel, Open Pi, the taskpane prompt, model/status footer, and any feature under test.'
```

Required evidence:

- Excel window title/id and `isOnscreen=true`.
- Taskpane evidence (`Pi for Excel`, `Open Pi`, prompt/model/status or feature-specific text).
- Frontmost app before/after is the same non-Excel app.
- Optional screenshot saved with a helper script or `observe` image.

## Taskpane bridge commands

The taskpane bridge commands run in the real Excel taskpane process. Use them to **mutate scratch workbook state, read it back, and clean up**; passive read-only probes are only the baseline evidence.

```bash
TOKEN=$(cat "$TOKEN_FILE")

PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- status
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- officeProbe
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- readUsedRange
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- readRange '{"address":"Sheet1!A1:B5"}'
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- listCharts
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- configureProxy \
  '{"enabled":true,"url":"https://localhost:3003"}'
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- selectModel \
  '{"provider":"openai-codex","modelId":"gpt-5.6-sol"}'
```

### Controlled write smoke

Run this whenever the change affects workbook IO, Office.js host behavior, recovery, chart/range tools, or anything where browser-only tests are too weak:

```bash
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- workbookWriteProbe \
  '{"sheetName":"_pi_background_verify","keepSheet":false}'
```

`workbookWriteProbe` creates or reuses a scratch sheet, writes a marker + numeric inputs + formula, reads the resulting range, then deletes the created scratch sheet or restores the previous scratch range. Its output is the minimum proof that the real hidden taskpane can write to and read from the real workbook without foregrounding Excel.

For custom setup/assert/cleanup:

```bash
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- writeRange \
  '{"address":"Sheet1!A1:B2","values":[["pi background smoke",2],["sum",3]]}'
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- readRange '{"address":"Sheet1!A1:B2"}'
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- clearRange '{"address":"Sheet1!A1:B2","applyTo":"contents"}'
```

For acceptance, submit an actual prompt through the sidebar. Use a scratch sheet name that does not already exist. For example, select a currently approved model, then ask it to create the sheet and write the formulas:

```bash
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- selectModel \
  '{"provider":"openai-codex","modelId":"gpt-5.6-sol"}'
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- submitPrompt \
  '{"text":"Use workbook tools to create a new sheet _pi_e2e_smoke; stop if it exists. Write 2 and 3 into A1:A2, write =SUM(A1:A2) into A3, and read back the values and formula. Leave the sheet for independent inspection. Do not change other sheets.","waitForIdle":true,"timeoutMs":120000}' --timeout 130000
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- readRange '{"address":"_pi_e2e_smoke!A1:A3"}'
```

Assert that A3 is 5 and retains its formula. Ask the model to delete only the scratch sheet, then use `officeProbe` to confirm it is gone. Adapt the prompt to exercise the changed feature; this arithmetic example does not cover every change.

The bridge's `submitPrompt` uses the sidebar send path. Do not send `/new` or `/name` through it expecting slash-command dispatch: these can be sent literally to the model. Create a test chat through the actual UI, verify its identity and idle state, then select the model. A restored session label does not prove which model answered; check `lastAssistant.model` and `stopReason` after the real prompt.

Use the outputs as verification artifacts:

- `status`: proves taskpane origin, hidden/background visibility, Office/Excel globals, active runtime/model, input state.
- `officeProbe`: proves Office.js can read the real workbook from the hidden taskpane.
- `workbookWriteProbe`: proves the hidden taskpane can perform reversible real workbook writes and read back formula results.
- `writeRange` / `clearRange`: deterministic setup and cleanup for feature-specific smoke tests.
- `configureProxy`: explicitly enables/disables the app's configured proxy for transport-specific real-host checks.
- `selectModel`: opens the real model selector, filters it, clicks the exact provider/model row, and verifies that model became active.
- `submitPrompt`: exercises the real app prompt → runtime/model/tool loop from the hidden taskpane.
- `readRange` / `readUsedRange`: verify workbook contents changed as expected.
- `listCharts`: verify chart creation/update/delete metadata.

For adversarial verification, wrap mutating and read-back commands with a frontmost check using `pi-computer-use` helper or strict AX `list_apps`. Completion evidence must include `frontmostSameApp=true` and `frontmostSameWindow=true` (or equivalent before/after app/window IDs).

## Safety rules

- Keep the bridge loopback-only and tokened. Never run it without `PI_BACKGROUND_VERIFY_TOKEN`.
- Never commit or print the token in docs/logs intended for sharing.
- The bridge is dev-only: the taskpane only connects when Vite injects `VITE_PI_BACKGROUND_VERIFY_URL` and `VITE_PI_BACKGROUND_VERIFY_TOKEN`.
- Prefer controlled, reversible writes over passive observation when workbook behavior is under test. Use scratch sheets/ranges, capture before/after output, and clean up or restore state.
- For app-level verification, use the normal product path (`submitPrompt` or the specific UI/tool path under test) and then assert the resulting workbook state via bridge read-back.
- If `pi-computer-use` offers raw pointer/keyboard/focus fallback, refuse it for background verification.
- Clean up when done:

```bash
kill "$(cat /tmp/pi-background-verify-bridge-server.pid)" 2>/dev/null || true
kill "$(cat /tmp/pi-for-excel-vite-background-verify.pid)" 2>/dev/null || true
rm -f "$TOKEN_FILE"
```

## Known limitations

- Strict AX can observe and press some semantic controls, but WebKit text-area `setText` may fail. Use the bridge for verification rather than typing prompts into the taskpane.
- Unsaved workbooks may have `workbookContext.workbookId=null`; use sheet/range/chart evidence instead.
- This is a background verification lane, not a replacement for full foreground/manual release smoke on a dedicated host when raw GUI interaction is the behavior under test.
