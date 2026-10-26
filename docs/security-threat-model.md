# Security threat model (v1)

This document summarizes what Pi for Excel stores, where data flows, and key trust boundaries.

## Scope

- Excel taskpane app in Office webviews (WKWebView/WebView2/browser)
- Hosted static build + optional local helper services (CORS proxy, tmux bridge, Python bridge)
- Credential flows (API keys + browser OAuth)
- Extension runtime model (host vs sandbox)

## Sensitive data

- Provider API keys (IndexedDB `ProviderKeysStore`)
- OAuth credentials (IndexedDB settings `oauth.<provider>`)
- Workbook contents read by tools
- Conversation/session history (IndexedDB)

## Storage model

- API keys: IndexedDB store via vendored first-party storage backend (`src/storage/local/`)
- OAuth credentials: IndexedDB settings (`oauth.<provider>`)
- Sessions/settings: IndexedDB

### User controls

- `/login` can add/replace/disconnect providers
- Disconnect removes provider key and OAuth credentials for that provider
- `/settings` includes API key + proxy configuration

## Network model

Taskpane communicates with:
- Office JS CDN (`appsforoffice.microsoft.com`)
- configured model/OAuth providers
- optional local HTTPS proxy (`https://localhost:<port>`)
- optional local bridge services (tmux / Python)

Hosted taskpane is protected with CSP in `vercel.json` (scripts/styles/fonts/connect constrained to required endpoints).

## Trust boundaries

1. **Taskpane webview** (untrusted workbook/model text can enter UI)
2. **Local helper services** (proxy/bridges are separate trust boundaries)
3. **Remote providers** (LLM + OAuth endpoints)
4. **Extension runtime boundary** (host runtime vs sandbox iframe runtime)

## Main threats and current controls

### 1) XSS/content injection in markdown/UI
- Marked safety patch blocks unsafe link protocols
- Markdown images are rendered as links (no automatic `<img>` fetch)
- Dynamic HTML sinks use escaping helpers where needed
- CSP reduces script/connect exfil paths

### 2) Token leakage via browser storage/logs
- OAuth credentials are stored in IndexedDB settings (no legacy localStorage fallback)
- No intentional token logging in auth restore/proxy paths
- Provider disconnect clears both API key and OAuth credentials

### 3) Local proxy / bridge abuse (CORS/SSRF/local attack surface)
- Loopback client requirement
- Allowed-origin CORS allowlist
- Strict target filtering/allowlists for proxy traffic
- Optional bearer-token auth on tmux/python bridge POST endpoints
- Bounded payload sizes + execution timeouts in helper servers

### 4) Extension code execution risks
- Remote `http(s)` extension URLs are blocked by default (`/experimental on remote-extension-urls` required). The opt-in is checked at install and at every load, in both host and sandbox runtimes (the sandbox imports module URLs itself, so the host checks before creating it)
- Untrusted extension sources (inline code + remote URL) run in sandbox iframe runtime by default
- Rollback kill switch exists for maintainers (`/experimental on extension-sandbox-rollback`) and should be temporary only
- Capability permissions are persisted per extension (`extensions.registry.v2`) and always enforced; the Plugins page toggles take effect on the next load (enabled extensions reload at once)
- Pi installing, enabling or uninstalling an extension asks for approval in every execution mode; install shows the code and any extra capabilities requested, which are the only grants beyond the pasted-code defaults
- Pi's Files writes/deletes ask for approval in Confirm mode, and in Auto mode whenever they would delete or overwrite a file Pi did not create in that chat

## Known limitations

- IndexedDB is not an XSS boundary; same-origin script execution can read stored credentials.
- Built-in/local-module extensions are trusted and run in host runtime.
- Sandbox runtime intentionally limits API surface (for example, no raw `api.agent` in sandbox).
- Host-specific CSP behavior still needs smoke testing across Excel macOS/Windows/Web.
- Tool-argument schema validation is intentionally disabled in Office builds (Ajv uses runtime code generation blocked by Office CSP, so the browser build aliases Ajv to stubs).

## Operational guidance

- Prefer localhost HTTPS proxy only; remote proxies can observe prompts/tokens.
- Keep dependencies updated (CI + Dependabot + audit checks).
- When adding new outbound endpoints, update CSP + proxy/docs/tests in the same PR.
