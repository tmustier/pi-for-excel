---
name: excel-background-verification
description: Verify pi-for-excel work against a real local Excel desktop host while keeping Excel in the background. Use after changes to workbook tools, recovery, taskpane UI, Office.js host behavior, chart/image paths, or any adversarial verification where browser-only tests are insufficient.
---

# Excel Background Verification

Use this skill to verify work in **real Excel with the real pi-for-excel taskpane** without taking over the user's foreground app.

The workflow has two channels:

1. **Computer-use observation** (`pi-computer-use`, strict AX mode) verifies the Excel window/taskpane visually and semantically while asserting the foreground app does not change.
2. **Taskpane background bridge** lets the real taskpane run Office.js probes, controlled scratch writes, and optional prompt submissions from inside Excel via a tokened loopback HTTPS server.

Do not use raw keyboard/mouse/coordinate actions in this workflow. If a step requires foreground focus or raw input, stop and report that the background lane cannot verify that path.

## Prerequisites

- Excel has the dev manifest sideloaded, pointing to `https://localhost:3141/src/taskpane.html`.
- `cert.pem` and `key.pem` exist in the repo root and are trusted by the local WebView.
- `pi-computer-use.app` has macOS Accessibility + Screen Recording permissions.
- Run from the repo root.

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

If `clients` is empty, reload the taskpane without foregrounding Excel: use `pi-computer-use` in strict AX mode to press the Pi taskpane close button, then press the ribbon **Open Pi** button. Do not use raw input fallbacks.

If `clients` lists more than one taskpane, pass `--clientId <id-from-health>` to mutating `background:verify:command` calls so writes target the intended workbook. The server refuses untargeted commands when multiple live clients are connected.

## Computer-use observe lane

Run Pi in strict AX mode and allow only observation/search tools unless a semantic AX button press is explicitly required:

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

When the product path itself is under test, submit an actual prompt through the sidebar instead of only calling Office.js helpers:

```bash
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- submitPrompt \
  '{"text":"Write SMOKE into A1, then report exactly what changed.","waitForIdle":true,"timeoutMs":120000}'
PI_BACKGROUND_VERIFY_TOKEN="$TOKEN" PI_BACKGROUND_VERIFY_HOST=localhost npm run background:verify:command -- readRange '{"address":"Sheet1!A1"}'
```

Use the outputs as verification artifacts:

- `status`: proves taskpane origin, hidden/background visibility, Office/Excel globals, active runtime/model, input state.
- `officeProbe`: proves Office.js can read the real workbook from the hidden taskpane.
- `workbookWriteProbe`: proves the hidden taskpane can perform reversible real workbook writes and read back formula results.
- `writeRange` / `clearRange`: deterministic setup and cleanup for feature-specific smoke tests.
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
