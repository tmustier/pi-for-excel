# pi-for-excel-tmux-bridge

Local HTTPS tmux bridge helper for Pi for Excel.

## Usage

```bash
npx pi-for-excel-tmux-bridge
```

This command:

1. Ensures `mkcert` exists (installs via Homebrew on macOS if missing)
2. Creates certificates in `~/.pi-for-excel/certs/` when needed
3. Starts the bridge at `https://localhost:3341`
4. Runs a real `tmux` backend

`tmux` must be installed and available on `PATH`; otherwise the bridge exits with an install hint.

Optional assisted install (macOS/Homebrew):

```bash
npx pi-for-excel-tmux-bridge --install-missing
```

This installs missing `tmux` before starting the bridge.

Then in Pi for Excel:

1. The default tmux bridge URL is already `https://localhost:3341`
2. (Optional) set `/experimental tmux-bridge-url <url>` to use a non-default URL
3. (Optional) run `/experimental tmux-bridge-token <token>` if you set `TMUX_BRIDGE_TOKEN`

## Publishing (maintainers)

Package source lives in `pkg/tmux-bridge/`.

Before packing/publishing, `prepack` copies runtime files from repo root:

- `scripts/tmux-bridge-server.mjs`

Publish from this directory:

```bash
cd pkg/tmux-bridge
npm publish
```
