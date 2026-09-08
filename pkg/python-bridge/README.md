# pi-for-excel-python-bridge

Local HTTPS Python / LibreOffice bridge helper for Pi for Excel.

## Usage

```bash
npx pi-for-excel-python-bridge
```

This command:

1. Ensures `mkcert` exists (installs via Homebrew on macOS if missing)
2. Creates certificates in `~/.pi-for-excel/certs/` when needed
3. Starts the bridge at `https://localhost:3340`
4. Runs real local Python and LibreOffice commands

`python3` must be on `PATH`. LibreOffice (`soffice` / `libreoffice`) is optional for Python execution but required for `libreoffice_convert`.

Optional assisted install (macOS/Homebrew):

```bash
npx pi-for-excel-python-bridge --install-missing
```

This installs missing `python3` and/or LibreOffice before starting the bridge.

Then in Pi for Excel:

1. The default Python bridge URL is already `https://localhost:3340`
2. (Optional) set `/experimental python-bridge-url <url>` to use a non-default URL
3. (Optional) run `/experimental python-bridge-token <token>` if you set `PYTHON_BRIDGE_TOKEN`

## Publishing (maintainers)

Package source lives in `pkg/python-bridge/`.

Before packing/publishing, `prepack` copies runtime files from repo root:

- `scripts/python-bridge-server.mjs`

Publish from this directory:

```bash
cd pkg/python-bridge
npm publish
```
