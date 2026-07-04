# Smoke Run — macOS post-merge pass (pi-stack 0.80.3 + TS6 + marked 18 + central proxy + portless)

- Date: 2026-07-04
- Commit: `ebe4f27` (main, after #581/#583/#584/#587/#588/#589/#570/#590/#593/#596/#597/#594)
- Environment: macOS, Excel Desktop (real host) + Chromium-hosted taskpane via agent-browser
- Checklist source: `docs/release-smoke-test-checklist.md`
- Operator: agent-driven local run

## Scope

Post-merge verification of the 2026-07-04 dependency/feature batch before the next
release tag: pi-stack `@earendil-works/*@0.80.3` (`/compat` migration), TypeScript 6,
`marked` 18, ChatGPT/Codex credential reconciliation (#587/#597), central proxy config +
provider allowlist (#593), portless opt-in dev flow (#594).

## Preflight (PRE-1)

All run at repo root on `ebe4f27`:

1. `npm run check` — pass
2. `npm run build` — pass
3. `npm run test:models` — 47/47 pass
4. `npm run test:context` — 665/665 pass
5. `npm run test:security` — 83/83 pass
6. `npm run test:manifest` — 15/15 pass

## Browser-hosted taskpane (dev server `https://localhost:3000`, agent-browser)

- Taskpane boots outside Excel via Office.js fallback; no console/page errors.
  (First load raced a vite dependency re-optimization reload after `npm install`;
  clean on reload — not a product issue.)
- Welcome + proxy-down banner shows **local-build copy** (`npx pi-for-excel-proxy`
  helper steps), confirming `DEFAULT_PROXY_IS_REMOTE=false` copy branching from #593.
- Settings → Providers: full provider list rendered (allowlist default = all
  providers, fail-open confirmed); proxy section shows local recommended URL
  `https://localhost:3003` + install guide link.
- Live model call: prompt sent to GPT-5.5 via connected OpenAI (ChatGPT)
  credentials through the dev-server API proxy; streamed reply "OK"
  (↑8.1k ↓19 tokens) — validates pi-ai 0.80.3 `/compat` path in the real UI.
- Local CORS proxy (`node scripts/cors-proxy-server.mjs --https`): boots with
  post-#593 defaults — loopback-only client policy, default target-host
  allowlist, `GET /healthz` → `ok`. Enabling proxy in settings shows
  "Proxy connected at https://localhost:3003" and the banner clears.
- Session restore: full history + context usage restored after page reload.

## Excel Desktop host (sideloaded manifest → `https://localhost:3000`)

- Add-in taskpane loads inside Excel from the dev server; prior session history
  auto-restored in host (C-2-lite).
- End-to-end agent action: prompt "Write the word SMOKE into cell A1, then
  confirm what you wrote." → Thought + `Read Sheet1!A1` + `Edit Sheet1!A1 — 1 changed`
  tool cards rendered; reply "Wrote SMOKE into cell A1."
- Independent verification: AppleScript `get value of range "A1"` returned
  `SMOKE` — workbook mutation confirmed outside the agent's own claim.
- Model call in host used GPT-5.5 via ChatGPT subscription credentials
  (post-#587/#597 credential path).

## Checklist status snapshot

| ID | Status | Notes |
|---|---|---|
| PRE-1 | Pass | Full preflight + manifest suite on `ebe4f27`. |
| C-1 | Pass (lite) | In-host read + write of real workbook cell, verified externally. |
| C-2 | Pass (lite) | Session restore verified in browser host and Excel host. |
| C-3 | Not run | Single-cell edit reported "no backup"; full checkpoint/restore flow not exercised this pass. |
| C-4 | Not run | — |
| C-5 | Not run | — |
| P-1..P-4 | Not run | — |
| I-1 | Pass (existing sideload) | Manifest present, points at localhost:3000, add-in launches. |
| I-3 | Pass (lite) | Existing connected providers produced live responses in both hosts. |
| I-4 | Pass (lite) | Proxy start → settings enable → "Proxy connected" verified; proxy-down banner + remediation copy verified beforehand. |
| H-1 | Partial | Proxy-down state verified (banner + steps). Other failure injections not run. |
| H-2 | Not run | — |
| H-3 | Pass | `npm run test:security` 83/83; live proxy boot shows loopback-only client policy + strict default target allowlist. |
| H-4 | Not run | — |

## Observations / follow-ups

- `Edit Sheet1!A1 — 1 changed, no backup`: expected behavior, not a regression.
  The test workbook was never saved, so it had no workbook identity;
  `src/workbook/recovery-log.ts` intentionally skips snapshot capture when
  `workbookIdentity` is unavailable (see `resolveWorkbookIdentity` guard).
  Checkpoints apply to saved workbooks.
- Windows pass (I-2) still outstanding, as in prior runs.
