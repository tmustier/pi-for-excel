# Smoke Run — macOS Excel host H-1 operator checklist (template)

Use this focused runbook to close **#179 / H-1** in a real Excel Desktop host.

- Date: YYYY-MM-DD
- Commit: `git rev-parse --short HEAD`
- Environment: macOS version + Excel version/build + model/provider
- Checklist source: `docs/release-smoke-test-checklist.md`

## Goal

Validate all H-1 failure paths produce clear, recoverable UX:
- explicit error surfaced (tool card or assistant error block)
- actionable next step shown
- no frozen spinner / no blank taskpane

## Pre-run setup

1. Start from clean add-in state (Excel fully quit/reopen).
2. Sideload latest manifest and open Pi taskpane.
3. Ensure screenshot path exists, e.g. `~/Desktop/pi-h1-<date>/`.
4. In Pi:
   - open `/integrations`
   - enable `Allow external tools (web search / MCP)`
   - enable web search for this session
5. Record baseline screenshot (`h1-baseline.png`).

## Scenario matrix (execute top-to-bottom)

| ID | Scenario | Status (Pass/Fail/Blocked) | Evidence (screenshot/log) | Notes |
|---|---|---|---|---|
| H1-A | Wrong API key |  |  |  |
| H1-B | Expired OAuth token |  |  |  |
| H1-C | Proxy enabled but down |  |  |  |
| H1-D | Rate limit during streaming |  |  |  |
| H1-E | Network disconnect mid-stream |  |  |  |

---

## H1-A — Wrong API key

1. `/integrations` → Web Search provider = Serper (or your configured provider).
2. Set key to an intentionally invalid value (e.g. `bad-key-h1`).
3. Prompt: `Use web_search for 'EUR USD exchange rate' and show raw result.`

Expected:
- clear auth failure (`401`/`403`/Unauthorized)
- no hang; input becomes usable again

Capture:
- `h1-a-wrong-key.png`
- exact visible error string in notes

---

## H1-B — Expired OAuth token

Preferred (deterministic, localhost/dev manifest) approach:

1. Quit Excel.
2. Backup `~/.pi/agent/auth.json`.
3. For the OAuth provider under test (e.g. Anthropic), replace token fields with obviously invalid values while keeping valid JSON.
4. Reopen Excel + taskpane.
5. Send a provider-backed prompt.

Expected:
- explicit token-expired/reauth-required error (`401`, `invalid_grant`, or equivalent)
- remediation is clear (login again / refresh credentials)

Capture:
- `h1-b-expired-oauth.png`
- exact error text
- restore original `~/.pi/agent/auth.json` after test

Notes:
- `~/.pi/agent/auth.json` is the dev credential source used by `/__pi-auth`.
- If running production manifest (no `/__pi-auth`), use a deliberately expired/revoked OAuth session instead and record the method in notes.

---

## H1-C — Proxy enabled but down

1. `/settings` → Proxy: enable `Use CORS Proxy`.
2. Set proxy URL to a non-running endpoint (e.g. `https://localhost:3999`).
3. Prompt: `Use fetch_page on https://example.com and summarize in one line.`

Expected:
- deterministic transport error (`Failed to fetch`/connection refused)
- remediation guidance points to proxy availability/config

Capture:
- `h1-c-proxy-down-config.png`
- `h1-c-proxy-down-result.png`

---

## H1-D — Rate limit during streaming

Recommended method:

1. Use a provider/key known to be near/over quota (or a controlled test account).
2. Send a response-heavy prompt (forces longer streaming), e.g.:
   - `Write a 3000-word analysis with 20 bullet sections and detailed examples.`
3. Repeat once if needed to trigger quota/rate limiting.

Expected:
- rate-limit failure is explicit (`429`, `rate limit`, `too many requests`)
- stream stops gracefully; UI recovers without manual reload

Capture:
- `h1-d-rate-limit.png`
- error text + whether failure occurred mid-response vs at start

If true mid-stream 429 is not reproducible, mark `Blocked` and document attempted method + observed fallback behavior.

---

## H1-E — Network disconnect mid-stream

1. Start a long-running prompt (same as H1-D is fine).
2. While response is actively streaming, disable network (Wi-Fi off / disconnect adapter).
3. Wait for failure to surface.
4. Re-enable network.

Expected:
- clear network interruption error (not silent stop)
- spinner clears; input usable; no blank taskpane
- user can retry after network returns

Capture:
- `h1-e-network-drop.png`
- note timing (seconds after stream start when disconnected)

---

## Exit criteria

- [ ] H1-A..H1-E are all `Pass`, or explicitly `Blocked` with concrete blocker + owner
- [ ] Evidence links are present for each scenario
- [ ] Top-level checklist row `H-1` updated in `docs/release-smoke-test-checklist.md` evidence table/run log
