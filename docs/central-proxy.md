# Org-hosted central CORS proxy

**Audience:** IT admins / platform teams rolling out Pi for Excel across an organisation.

By default, Pi for Excel expects each user to run the CORS proxy locally (`https://localhost:3003`, see [install.md](./install.md#oauth-logins-and-cors-proxy)). That requires Node.js on every machine. For managed rollouts you can instead run **one proxy on a central server** and build the add-in so it points there by default.

This guide covers both halves:

1. Deploying `scripts/cors-proxy-server.mjs` centrally
2. Building an org-configured add-in (default proxy URL, provider allowlist, CSP, manifest)

> **Security model up front:** the proxy forwards whatever `Authorization`/API-key headers the caller provides to *allowlisted target hosts only*. It does not store credentials. Restricting **who can reach the proxy** is your job (network segmentation + `ALLOWED_CLIENT_CIDRS`); restricting **where it can forward** is `ALLOWED_TARGET_HOSTS`. There is currently no built-in client auth token — do not expose the proxy to the public internet.

---

## 1) Deploy the proxy centrally

The proxy is a single-file Node server with no dependencies:

```bash
node scripts/cors-proxy-server.mjs --https
```

### Environment variables

| Variable | Default | Purpose |
|---|---|---|
| `HOST` | `localhost` (https) / `127.0.0.1` (http) | Bind address. Use your server's interface address (or `0.0.0.0` behind a firewall). |
| `PORT` | `3003` | Listen port. |
| `HTTPS=1` / `--https` | off | Serve TLS directly. Recommended unless you terminate TLS at a reverse proxy. |
| `TLS_KEY_PATH` / `TLS_CERT_PATH` | `./key.pem` / `./cert.pem` | Paths to your org-issued TLS key/cert (e.g. `*.example.com`). |
| `ALLOWED_CLIENT_CIDRS` | *(unset — loopback only)* | Comma-separated IPv4 CIDRs (or bare IPs) allowed as clients, e.g. `10.96.0.0/13,192.168.1.5`. Loopback is always allowed. Invalid entries (including `/0`) are fatal at startup — fail closed. IPv6 client ranges are not supported. |
| `ALLOWED_ORIGINS` | dev + hosted origins | Comma-separated browser origins allowed by CORS. Set to the origin serving *your* add-in build, e.g. `https://pi-excel.example.com`. |
| `ALLOWED_TARGET_HOSTS` | built-in provider allowlist | Comma-separated hosts the proxy may forward to. For a locked-down org, set exactly the hosts of your approved providers, e.g. `api.deepseek.com` or an internal gateway host. **When explicitly set, this allowlist also applies to loopback/private targets** — the override flags below cannot bypass it. |
| `ALLOW_PRIVATE_TARGETS=1` | off | Required only if a target is a private/internal address (e.g. an on-prem LLM gateway at `10.x.x.x`). With an explicit `ALLOWED_TARGET_HOSTS`, the private IP or gateway hostname must **also** be listed there (e.g. `ALLOWED_TARGET_HOSTS=api.deepseek.com,10.97.193.77`) — other private addresses stay blocked. Without an explicit allowlist (legacy local behavior), this flag allows *any* private target: never run that combination on a shared proxy. |
| `STRICT_TARGET_RESOLUTION=1` | off | Reject targets whose DNS doesn't resolve. Recommended on servers. |

Do **not** set `ALLOW_ALL_TARGET_HOSTS=1` on a shared proxy.

### Example: systemd-style launch

```bash
HOST=0.0.0.0 \
PORT=3003 \
HTTPS=1 \
TLS_KEY_PATH=/etc/pi-proxy/wildcard.example.com.key \
TLS_CERT_PATH=/etc/pi-proxy/wildcard.example.com.pem \
ALLOWED_CLIENT_CIDRS=10.96.0.0/13 \
ALLOWED_ORIGINS=https://pi-excel.example.com \
ALLOWED_TARGET_HOSTS=api.deepseek.com,internal-llm.example.com \
ALLOW_PRIVATE_TARGETS=1 \
STRICT_TARGET_RESOLUTION=1 \
node scripts/cors-proxy-server.mjs
```

(`ALLOW_PRIVATE_TARGETS=1` here is only needed because `internal-llm.example.com` resolves to a private address — and it stays constrained to the hosts listed in `ALLOWED_TARGET_HOSTS`.)

At startup the proxy logs its effective client, origin, and target policies — check them.

### Health checks / reverse proxies

- `GET /healthz` returns `200 ok` without requiring an `Origin` header (subject to the client-address check). Point load-balancer health checks here.
- All *proxying* requests still require an allowlisted `Origin`. If you front the proxy with nginx, make sure it passes the `Origin` header through unchanged (nginx does by default; don't override it).
- If you terminate TLS at the reverse proxy, note the client-address check sees the reverse proxy's address — restrict at the network layer accordingly (the proxy does not trust `X-Forwarded-For`).

## 2) Build an org-configured add-in

Central deployments self-host the static build (fork or CI job). Two build-time env vars configure the client:

| Variable | Example | Effect |
|---|---|---|
| `VITE_PI_DEFAULT_PROXY_URL` | `https://pi-proxy.example.com:3003` | Default proxy URL baked into the build (users can still change it in `/settings`). Must be `https://`. |
| `VITE_PI_ALLOWED_PROVIDERS` | `deepseek,openai` | Only show these provider ids in the connect UI. Ids match `ALL_PROVIDERS` in `src/ui/provider-login.ts` (e.g. `anthropic`, `openai-codex`, `openai`, `google`, `deepseek`, `mistral`, `groq`, `xai`, ...). **UI filter only** — pair it with `ALLOWED_TARGET_HOSTS` on the proxy for actual enforcement. |

```bash
VITE_PI_DEFAULT_PROXY_URL=https://pi-proxy.example.com:3003 \
VITE_PI_ALLOWED_PROVIDERS=deepseek,openai \
npm run build
```

### CSP (`vercel.json` or your host's headers)

The hosted build's `Content-Security-Policy` only allows proxy connections to `https://localhost:*`. Add your proxy origin to `connect-src`, e.g.:

```
... https://localhost:* https://127.0.0.1:* https://pi-proxy.example.com:3003; ...
```

If your org restricts providers, you can also *remove* unused provider hosts from `connect-src` for defence in depth.

### Manifest

Generate a manifest pointing at your hosting origin:

```bash
ADDIN_BASE_URL="https://pi-excel.example.com" OUT=manifest.org.xml node scripts/generate-manifest.mjs
```

Distribute via a [network-share catalog (Windows)](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) or centralized deployment.

## 3) Verify the rollout

From a user machine (inside the allowed network):

```bash
curl https://pi-proxy.example.com:3003/healthz
# → ok

curl -H "Origin: https://pi-excel.example.com" \
  "https://pi-proxy.example.com:3003/?url=https%3A%2F%2Fapi.deepseek.com%2F"
# → provider response (or its 401), not a 403 from the proxy
```

Then in Excel: open the add-in → `/settings` → Proxy should show your org URL and a green reachability state; the connect list should show only approved providers.

## Limitations / future work

- **No client auth token yet** — client restriction is network + CIDR based. Tracked in [#595](https://github.com/tmustier/pi-for-excel/issues/595).
- `ALLOWED_CLIENT_CIDRS` is IPv4-only.
- The official hosted build (`pi-for-excel.vercel.app`) cannot use an org proxy because of its CSP; central-proxy setups must self-host the static build.
