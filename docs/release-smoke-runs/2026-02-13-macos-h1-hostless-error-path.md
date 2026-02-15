# Smoke Run — macOS hostless H-1 error-path focus (taskpane browser harness)

- Date: 2026-02-13
- Commit: `62bff943701068e85b042e2268c721fd3d4ebe31`
- Environment: macOS CLI + Playwright (`agent-browser`) against `https://localhost:3000/src/taskpane.html`
- Checklist source: `docs/release-smoke-test-checklist.md`
- Scope note: **hostless** browser harness only (not desktop Excel host). This validates error rendering/UX behavior in the taskpane UI, not workbook-host integration semantics.

## Setup used

- `npm ci`
- `npm run dev` (local Vite taskpane)
- Opened taskpane with `agent-browser`
- Opened Integrations overlay and enabled:
  - `Allow external tools (web search / MCP)`
  - `Web Search` session toggle
  - `MCP` session toggle
- Configured bad Serper key (`bad-serper-key`) to force auth failure.
- Enabled proxy in settings and pointed proxy URL to `https://localhost:3999` (intentionally down) to force transport failure.

## H-1 matrix evidence

| Scenario | Status | Evidence | Notes |
|---|---|---|---|
| Wrong API key | Pass | `/tmp/h1-body.html:99` (`Error: Serper.dev (default) search request failed (403): {"message":"Unauthorized.","statusCode":403}`), screenshot `/tmp/h1-wrong-key-result.png` | Deterministic tool-card error text surfaced; no blank UI. |
| Expired OAuth token | Pass | `/tmp/h1-network-disconnect-body.html:319` (`Error: 401 ... OAuth token has expired`), screenshot `/tmp/h1-network-disconnect-result.png` | Taskpane rendered explicit auth remediation context, not spinner hang. |
| Proxy enabled but down | Pass | `/tmp/h1-proxy-down-body.html:208` (`Error: Failed to fetch`), screenshot `/tmp/h1-proxy-configured-down.png`, `/tmp/h1-proxy-down-result.png` | Failure surfaced as readable fetch error after proxy misconfiguration/down simulation. |
| Rate limit mid-stream | Blocked | Attempted `fetch_page https://httpstat.us/429`; observed transport-level `Failed to fetch` at `/tmp/h1-rate-limit-body.html:282` | Could not reliably produce true streamed 429 path in this hostless setup after proxy/auth failures; keep as host-required/manual follow-up. |
| Network disconnect mid-stream | Blocked | Connection-level failures observed (`Failed to fetch`) but not true in-flight stream disconnect | Mid-stream disconnect remains better covered by automated `tests/error-path-matrix.test.ts` + host/manual check. |

## UX behavior check

Observed after each reproduced error:
- explicit error card/message rendered in transcript
- input remained available (no persistent "Stop"/working lock)
- no blank-screen failure

## Blockers / follow-ups

1. Still need **desktop Excel host** run to fully close #179 H-1 in the intended environment.
2. Remaining H-1 cases to validate in-host:
   - true streamed rate-limit interruption
   - true mid-stream network drop while model/tool call is active

## Related artifacts

- `/tmp/h1-wrong-key-integration.png`
- `/tmp/h1-wrong-key-result.png`
- `/tmp/h1-proxy-configured-down.png`
- `/tmp/h1-proxy-down-result.png`
- `/tmp/h1-rate-limit-result.png`
- `/tmp/h1-network-disconnect-result.png`
- `/tmp/h1-body.html`
- `/tmp/h1-proxy-down-body.html`
- `/tmp/h1-rate-limit-body.html`
- `/tmp/h1-network-disconnect-body.html`
- `/tmp/h1-console.json`
