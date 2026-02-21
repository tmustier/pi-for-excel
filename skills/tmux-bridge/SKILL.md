---
name: tmux-bridge
description: Local terminal access via the tmux bridge. Use when the user asks about running shell commands, setting up the tmux bridge, or troubleshooting terminal connectivity.
compatibility: Requires a local tmux bridge process running on the user's machine.
metadata:
  tool-name: tmux
  docs: docs/tmux-bridge-contract.md
---

# Tmux Bridge

The tmux bridge gives Pi access to a real local terminal on the user's machine. The `tmux` tool is always registered — it just needs the bridge process running locally to work.

## What it does

When the bridge is running, the `tmux` tool can:
- **list_sessions** — see active tmux sessions
- **create_session** — start a new shell session (optionally in a specific directory)
- **send_keys** — type commands into a session
- **capture_pane** — read terminal output
- **send_and_capture** — send a command and wait for output in one call
- **kill_session** — close a session

## Running `pi` (or other local CLIs) via tmux

The tmux pane is a normal local shell. If `pi` is installed, you can invoke it directly with `send_keys`/`send_and_capture` text like any other command.

Recommended flow:
1. `list_sessions` then `create_session` (or reuse an existing session)
2. Optional one-time check: `command -v pi`
3. Send the `pi ...` command
4. Monitor output with `capture_pane`

For long-running jobs, avoid rapid repeated captures. Prefer:
- `capture_pane` with `wait_ms` (for example 15000-30000), or
- `send_and_capture` with `wait_for` + `timeout_ms` when you know a completion pattern.

## How to set it up

### 1. Start the bridge

The bridge is a local HTTPS server. Run it from a terminal:

```bash
npx pi-for-excel-tmux-bridge
```

This defaults to **real tmux mode** on `https://localhost:3341`.

Options:
- `--install-missing` — auto-install tmux via Homebrew (macOS)
- `TMUX_BRIDGE_MODE=stub` — safe simulated mode (no real shell execution)
- `TMUX_BRIDGE_TOKEN=your-secret` — require auth token

### 2. Configure in Pi (usually not needed)

The default bridge URL (`https://localhost:3341`) works automatically — no configuration required. If you need a custom URL or auth token:

```
/experimental tmux-bridge-url <url>
/experimental tmux-bridge-token <token>
/experimental tmux-status
```

### 3. Accept the local HTTPS certificate

The bridge uses a self-signed cert. You may need to visit `https://localhost:3341` in your browser once and accept it.

## When the bridge is not running

The `tmux` tool stays registered but returns an error if the bridge is unreachable. Python tools (`python_run`, `python_transform_range`) still work via the in-browser Pyodide fallback — tmux is the only tool that strictly requires its bridge.

## Security

- Loopback-only (localhost)
- Origin allowlist
- Optional bearer token auth
- Session names and key tokens are validated
- No shell interpolation (argv-based tmux calls)

## Troubleshooting

- **"bridge URL is unavailable"** — the bridge process isn't running. Start it with `npx pi-for-excel-tmux-bridge`.
- **"timed out"** — the bridge is running but the command took too long. Default timeout is 15s; use `timeout_ms` for longer operations.
- **CORS/cert errors** — visit the bridge URL directly in your browser and accept the certificate.
