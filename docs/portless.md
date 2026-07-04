# Dev server behind portless (opt-in)

**Status:** opt-in recipe. The default dev flow (`npm run dev` on
`https://localhost:3141` with mkcert certs) is unchanged and remains what the
manifest, sideload docs, and CI assume. Use this page only if you want the
[portless](https://portless.sh) flavor: a stable named URL
(`https://pi-excel.localhost`) with no mkcert step and no fixed port.

Tracking issue: [#586](https://github.com/tmustier/pi-for-excel/issues/586).

## What portless does

Portless runs a local HTTPS reverse proxy on port 443. It generates and
trusts its own local CA, assigns your app a random port via `PORT` /
an injected `--port` flag, and proxies `https://pi-excel.localhost` → Vite.

That means in this mode:

- **No mkcert** — TLS terminates at the portless proxy; Vite serves plain
  HTTP on loopback. `key.pem` / `cert.pem` are not used.
- **No fixed port** — portless picks a free port (4000–4999) per run.
- **HMR still works** — `vite.config.ts` routes the HMR websocket through
  the proxy (`wss://pi-excel.localhost`).

## Requirements

- portless is a pinned `devDependency`, so `npm install` provides it — no
  global install needed.
- portless itself requires **Node 24+** (the repo otherwise supports
  `^22.19.0 || >=24`; on Node 22 you'll see an `EBADENGINE` warning at
  install time and `npm run dev:portless` won't work — the default flow
  still does).
- First run: portless generates a local CA, trusts it, and binds port 443
  (auto-elevates with `sudo` on macOS/Linux).

## Run

```bash
npm run dev:portless   # = portless pi-excel vite → https://pi-excel.localhost
```

How the Vite side switches over (see `vite.config.ts`):

- portless injects `PORTLESS_URL` into the child process; when it is set (or
  `DEV_HOST=<hostname>` is set explicitly, e.g. for another local HTTPS
  proxy), Vite disables local HTTPS, binds loopback only, takes the
  portless-assigned port, and points HMR at the proxy.
- With neither env var set, nothing changes: `npm run dev` behaves exactly
  as before.

## Sideload manifest

`manifest.xml` stays pinned to `https://localhost:3141`. Generate a
dev-proxy manifest instead (never committed; it's gitignored):

```bash
npm run manifest:dev                          # → manifest.dev.xml (https://pi-excel.localhost)
npm run manifest:dev -- my-addin.localhost    # custom hostname
node scripts/validate-manifest.mjs manifest.dev.xml
```

Sideload `manifest.dev.xml` the same way as `manifest.xml` (see README).
macOS:

```bash
cp manifest.dev.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

Note: the add-in Id is unchanged, so the dev-proxy manifest **replaces** the
default `localhost:3141` sideload — you run one or the other, not both. To
switch back, sideload `manifest.xml` again.

## Bridge / proxy servers (deliberate opt-in)

The local CORS proxy and tmux/python bridges enforce strict origin
allowlists pinned to `https://localhost:3141` (plus the hosted origin). The
portless origin is **deliberately not added by default** — extend the
allowlist explicitly via the existing `ALLOWED_ORIGINS` env var when you run
them:

```bash
ALLOWED_ORIGINS="https://localhost:3141,https://pi-for-excel.vercel.app,https://pi-excel.localhost" npm run proxy:https
ALLOWED_ORIGINS="https://localhost:3141,https://pi-for-excel.vercel.app,https://pi-excel.localhost" npm run tmux:bridge:https
ALLOWED_ORIGINS="https://localhost:3141,https://pi-for-excel.vercel.app,https://pi-excel.localhost" npm run python:bridge:https
```

Note `ALLOWED_ORIGINS` **replaces** the default list, so include every origin
you still want. Do not loosen the checked-in defaults — see
`docs/security-threat-model.md`.

The bridges' own HTTPS listeners still use the mkcert `key.pem`/`cert.pem`
files (they are separate servers on their own ports, not proxied by
portless). If you never ran mkcert, run the bridges in plain-HTTP loopback
mode (`npm run proxy`, etc.) or generate the certs as in the README.

## Caveats

- **Excel webview trust/resolution.** Browsers resolve `*.localhost` and
  trust the portless CA once installed in the system store; Excel's webview
  (WKWebView on macOS, WebView2 on Windows) uses the system resolver and
  cert store, so `portless trust` + its `/etc/hosts` syncing are what make
  the sideloaded add-in load. If the taskpane comes up blank, verify
  `https://pi-excel.localhost` loads in Safari (macOS) / Edge (Windows)
  first, and re-run `portless trust` after portless upgrades if needed.
- **Worktrees.** `portless run` prepends the branch name as a subdomain in
  linked git worktrees (e.g. `https://fix-ui.pi-excel.localhost`). If you
  use that, regenerate the manifest for the exact hostname
  (`npm run manifest:dev -- fix-ui.pi-excel.localhost`) and add it to
  `ALLOWED_ORIGINS` for any bridges.
- **portless is pre-1.0.** State directory format may change between
  releases; you may need to re-run `portless trust` after an upgrade.
