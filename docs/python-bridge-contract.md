# Python / LibreOffice bridge contract (v1)

Status: implemented as an optional **local helper** (native bridge). `python_run` and `python_transform_range` also support in-browser Pyodide fallback.

- Add-in adapters:
  - `src/tools/python-run.ts` (`python_run`)
  - `src/tools/libreoffice-convert.ts` (`libreoffice_convert`)
  - `src/tools/python-transform-range.ts` (`python_transform_range`)
- Local helper:
  - `scripts/python-bridge-server.mjs`

## Gate model

Bridge-backed tools remain registered (stable tool list / prompt caching).

Native bridge usage requires:

1. `python.bridge.url` is configured (`/experimental python-bridge-url <url>`)
2. bridge `GET /health` succeeds
3. user confirms the first Python/LibreOffice execution per configured bridge URL

Notes:
- `libreoffice_convert` is bridge-only and blocked when these checks fail.
- `python_run` / `python_transform_range` can still execute via Pyodide fallback when bridge checks fail.

Optional bearer auth:

- Set `PYTHON_BRIDGE_TOKEN` when starting the bridge
- Configure `/experimental python-bridge-token <token>`
- Stored key: `python.bridge.token`

---

## Local bridge quickstart

```bash
# Stub mode (safe default; deterministic fake responses)
npm run python:bridge:https

# Real local execution mode
PYTHON_BRIDGE_MODE=real npm run python:bridge:https
```

Then configure in the add-in:

```bash
/experimental python-bridge-url https://localhost:3340
# optional
/experimental python-bridge-token <token>
```

Or use the extensions UI: `/extensions` → **Local Python / LibreOffice bridge** → **Save URL**.

Bridge endpoints:

- `GET /health`
- `POST /v1/python-run`
- `POST /v1/libreoffice-convert`

High-level add-in helper tool:

- `python_transform_range` (read range → run Python → write transformed output)

---

## `POST /v1/python-run`

### Request

```json
{
  "code": "result = {'rows': [[1,2],[3,4]]}",
  "input_json": "{\"source\":\"Sheet1!A1:B2\"}",
  "timeout_ms": 10000
}
```

### Response

```json
{
  "ok": true,
  "action": "run_python",
  "exit_code": 0,
  "stdout": "...",
  "stderr": "...",
  "result_json": "{\"rows\":[[1,2],[3,4]]}",
  "truncated": false
}
```

Notes:

- In real mode, Python is executed locally (`python3` by default) with `-I` isolation flag.
- `input_json` is exposed to Python code as `input_data`.
- If code sets a `result` variable (JSON-serializable), bridge returns it as `result_json`.

---

## `POST /v1/libreoffice-convert`

### Request

```json
{
  "input_path": "/absolute/path/source.xlsx",
  "target_format": "pdf",
  "output_path": "/absolute/path/source.pdf",
  "overwrite": true,
  "timeout_ms": 60000
}
```

### Response

```json
{
  "ok": true,
  "action": "convert",
  "input_path": "/absolute/path/source.xlsx",
  "target_format": "pdf",
  "output_path": "/absolute/path/source.pdf",
  "bytes": 12345,
  "converter": "soffice"
}
```

Supported `target_format`: `csv`, `pdf`, `xlsx`.

Notes:

- In real mode the bridge shells out via argv (no shell interpolation):
  `soffice --headless --convert-to <format> --outdir <tmpDir> <input>`
- `input_path` / `output_path` must be absolute.

---

## Security posture

- Loopback-only client enforcement (`127.0.0.1` / `::1`)
- Origin allowlist (`ALLOWED_ORIGINS`; defaults include dev + hosted add-in origins)
- Optional bearer token auth for POST endpoints
- Request body + code/input/output size limits
- Timeout enforcement for Python and LibreOffice subprocesses

---

## Tool execution policy

`python_run` and `libreoffice_convert` are classified as **read / no workbook context impact**.

`python_transform_range` is classified as **mutate / content impact** because it writes transformed output into workbook cells.
