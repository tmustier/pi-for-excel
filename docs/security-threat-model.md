# Security threat model (v1)

This document summarizes what Pi for Excel stores, where data flows, and the key trust boundaries.

## Scope

- Excel taskpane app running in Office webviews (WKWebView/WebView2/browser)
- Hosted static build + optional local CORS proxy
- Credential flows (API keys + browser OAuth)

## Sensitive data

- Provider API keys (IndexedDB `ProviderKeysStore`)
- OAuth credentials (IndexedDB settings `oauth.<provider>`)
- Workbook contents read by tools
- Conversation/session history (IndexedDB)

## Storage model

- API keys: IndexedDB store via pi-web-ui storage backend
- OAuth credentials: IndexedDB settings (`oauth.<provider>`)
- Sessions/settings: IndexedDB

### User controls

- `/login` provider overlay can add/replace/disconnect providers
- Disconnect removes provider key and clears OAuth session for that provider
- `/settings` includes API key + proxy configuration

## Network model

Taskpane communicates with:
- Office JS CDN (`appsforoffice.microsoft.com`)
- configured model/OAuth providers
- optional local CORS proxy (`https://localhost:<port>`)

Hosted taskpane is protected with CSP in `vercel.json` (scripts/styles/fonts/connect sources constrained to required endpoints).

## Trust boundaries

1. **Taskpane webview** (untrusted workbook/model text can enter UI)
2. **Local helper services** (CORS proxy, dist server in dev/test)
3. **Remote providers** (LLM + OAuth endpoints)
4. **Extension loading boundary** (remote extension URL imports disabled by default)

## Main threats and current controls

### 1) XSS/content injection in markdown/UI
- Marked safety patch blocks unsafe link protocols
- Markdown images are rendered as links (no automatic `<img>` fetch)
- Dynamic HTML sinks use escaping helpers where needed
- CSP reduces script/connect exfil paths

### 2) Token leakage via browser storage/logs
- OAuth credentials are stored only in IndexedDB settings
- No intentional token logging in auth restore/proxy paths
- Provider disconnect clears both key and OAuth stored credentials

### 3) Local proxy abuse (CORS/SSRF)
- Loopback client requirement
- Allowed-origin CORS allowlist
- Default outbound host allowlist for known provider endpoints (plus path-scoped GitHub Enterprise OAuth/Copilot compatibility)
- Loopback/private target blocking by default (+ DNS-aware checks)
- Explicit opt-in overrides for advanced/local setups

### 4) Remote extension code execution
- `loadExtension()` blocks remote `http(s)` module URLs by default
- Protocol-relative URLs (`//host/...`) treated as remote
- Explicit unsafe local opt-in required for remote extension URL experiments

## Known limitations

- IndexedDB is not an XSS boundary; same-origin script execution can read stored credentials.
- Full extension sandbox/capability permissions are deferred until user-supplied extension distribution ships.
- Host-specific CSP behavior must continue to be smoke-tested in Excel macOS/Windows/Web.

## Operational guidance

- Prefer localhost HTTPS proxy only; remote proxies can observe prompts/tokens.
- Keep dependencies updated (CI + Dependabot + audit checks).
- When adding new outbound endpoints, update CSP + proxy/docs/tests in the same PR.
