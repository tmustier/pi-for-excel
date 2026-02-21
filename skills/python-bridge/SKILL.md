---
name: python-bridge
description: Native Python execution via the local Python bridge. Use when the user asks about running Python locally, setting up the Python bridge, LibreOffice conversion, or troubleshooting Python connectivity.
compatibility: Python tools always work via in-browser Pyodide. The native bridge is optional and requires a local bridge process running on the user's machine.
metadata:
  tool-name: python_run
  docs: docs/python-bridge-contract.md
---

# Python Bridge

The Python bridge gives Pi access to native Python on the user's machine. It is an opt-in, feature-flagged capability that upgrades the default in-browser Pyodide runtime.

## Pyodide (default) vs Native bridge

| | Pyodide (default) | Native bridge |
|---|---|---|
| Setup | None — works out of the box | Requires local bridge process |
| Packages | Pure-Python only (numpy, pandas, scipy via micropip) | Full ecosystem (C extensions, ML libs, etc.) |
| Filesystem | No local filesystem access | Full local filesystem |
| LibreOffice | Not available | Available (`libreoffice_convert` tool) |
| Long scripts | WebAssembly limits | No limits |

**Prefer Pyodide unless the task requires native-only capabilities.**

## How to set it up

### 1. Start the bridge

```bash
npx pi-for-excel-python-bridge
```

This defaults to **real execution mode** on `https://localhost:3340`.

Options:
- `--install-missing` — auto-install Python/LibreOffice via Homebrew (macOS)
- `PYTHON_BRIDGE_MODE=stub` — safe simulated mode
- `PYTHON_BRIDGE_TOKEN=your-secret` — require auth token
- `PYTHON_BRIDGE_PYTHON_BIN=python3.12` — specify Python binary

Requirements:
- `python3` must be on `PATH` (or set `PYTHON_BRIDGE_PYTHON_BIN`)
- LibreOffice (`soffice`) is optional — only needed for `libreoffice_convert`

### 2. Configure in Pi (optional)

The default URL (`https://localhost:3340`) works automatically. Override if needed:

```
/experimental python-bridge-url <url>
/experimental python-bridge-token <token>
```

Or use: `/extensions` → **Connections** → **Python bridge**

### 3. First execution

The first time Python runs through the native bridge, Pi will ask for explicit user confirmation (one-time per bridge URL).

## Tools

- **python_run** — execute a Python snippet, inspect stdout/stderr/result
- **python_transform_range** — read Excel range → run Python → write result back (single tool call)
- **libreoffice_convert** — convert files between formats (xlsx ↔ csv ↔ pdf) — bridge-only

## When the bridge is not running

`python_run` and `python_transform_range` fall back to Pyodide automatically. Only `libreoffice_convert` strictly requires the bridge.

## Security

- Loopback-only (localhost)
- Origin allowlist
- Optional bearer token auth
- Python runs with `-I` isolation flag
- Code/input/output size limits
- Timeout enforcement

## Troubleshooting

- **Falls back to Pyodide unexpectedly** — the bridge process isn't running. Start it with `npx pi-for-excel-python-bridge`.
- **Import errors on Pyodide** — the package likely has C extensions. Set up the native bridge.
- **LibreOffice convert fails** — ensure `soffice` is on PATH. Install with `brew install --cask libreoffice` (macOS).
- **CORS/cert errors** — visit `https://localhost:3340` in your browser and accept the certificate.
