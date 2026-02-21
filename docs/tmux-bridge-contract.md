# Tmux bridge contract (v1)

Status:
- Add-in adapter implemented in `src/tools/tmux.ts`
- Local bridge scaffold implemented in `scripts/tmux-bridge-server.mjs`

The bridge supports two modes:
- `tmux`: real tmux subprocess backend with guardrails
- `stub`: in-memory tmux simulation for development/testing (does not execute shell commands)

Notes:
- The one-command helper (`npx pi-for-excel-tmux-bridge`) defaults to `tmux` mode.
- The raw server script keeps `stub` as its default for local development/test usage.

## Availability and gating

The `tmux` tool remains registered (stable tool list / prompt caching), but execution is blocked unless all gates pass:

1. effective bridge URL is resolved (configured override via `/experimental tmux-bridge-url <url>`, else default `https://localhost:3341`)
2. bridge `GET /health` returns success

The gate is checked on each tool execution (defense in depth).

## Local bridge quickstart

```bash
# One-command setup (real tmux mode by default)
npx pi-for-excel-tmux-bridge

# Optional assisted dependency install (macOS/Homebrew)
npx pi-for-excel-tmux-bridge --install-missing

# Force safe simulated mode
TMUX_BRIDGE_MODE=stub npx pi-for-excel-tmux-bridge

# Source checkout alternative
npm run tmux:bridge:https
```

Real-mode requirement:

- `tmux` must be installed and discoverable on `PATH`
- `--install-missing` can install `tmux` on macOS/Homebrew

Then in the add-in:

```bash
# optional URL override (default is already https://localhost:3341)
/experimental tmux-bridge-url <url>
/experimental tmux-status
```

Optional auth token:

```bash
TMUX_BRIDGE_TOKEN=your-secret npx pi-for-excel-tmux-bridge
```

Store the same token for the tool adapter:

```bash
/experimental tmux-bridge-token <token>
```

(setting key: `tmux.bridge.token`)

## Endpoints

- `GET /health`
- `POST /v1/tmux`

Content-Type: `application/json`

Optional auth header when configured:
- `Authorization: Bearer <tmux.bridge.token>`

## Request schema

```json
{
  "action": "list_sessions | create_session | send_keys | capture_pane | send_and_capture | kill_session",
  "session": "optional session name",
  "cwd": "optional absolute working directory (create_session)",
  "text": "optional literal input (send_keys/send_and_capture)",
  "keys": ["optional key tokens, e.g. Enter, C-c"],
  "enter": true,
  "lines": 120,
  "wait_for": "optional regex string",
  "timeout_ms": 5000,
  "wait_ms": 15000,
  "join_wrapped": false
}
```

### Action requirements enforced by the add-in/bridge

- `list_sessions`: no required fields
- `create_session`: no required fields
- `capture_pane`: requires `session`
- `kill_session`: requires `session`
- `send_keys`: requires `session` + at least one of (`text`, `keys`, `enter=true`)
- `send_and_capture`: same as `send_keys`
- `wait_ms`: optional delay (0..120000ms) before capture for `capture_pane`/`send_and_capture`

Tip: `send_keys` sends input only. Use `capture_pane` or `send_and_capture` when you need terminal output. For long-running jobs, set `wait_ms` on capture calls (for example 15000-30000) instead of tight polling loops.

## Response schema

```json
{
  "ok": true,
  "action": "same action",
  "session": "optional resolved session",
  "sessions": ["optional list for list_sessions"],
  "output": "optional text output/capture",
  "error": "optional error string",
  "metadata": { "optional": "structured bridge metadata" }
}
```

Notes:
- Non-2xx HTTP responses are treated as errors by the adapter.
- `ok: false` is treated as an error by the adapter.
- Plain-text success responses are accepted as `output` fallback.

## Real tmux guardrails (implemented)

- Loopback client enforcement
- Origin allowlist enforcement (`ALLOWED_ORIGINS`)
- Optional bearer token auth (`TMUX_BRIDGE_TOKEN`)
- Session name validation (strict regex)
- Key token validation (strict regex)
- `cwd` must be absolute and an existing directory
- Bounded request size and input lengths
- Bounded `lines` and `timeout_ms`
- tmux calls executed via argv arrays (no shell interpolation)
- tmux launched with `-f /dev/null` and fixed socket path

## Tool behavior in workbook runtime

`tmux` is classified as read-only/non-workbook traffic in `execution-policy.ts`, so calls do not acquire workbook write locks or trigger workbook blueprint invalidation.
