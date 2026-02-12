# Pi for Excel

An open-source, multi-model AI sidebar add-in for Microsoft Excel — powered by [Pi](https://pi.dev).

**Bring your own key. Free. Open source.**

## What is this?

Pi for Excel puts an AI assistant directly in your Excel sidebar. It can read your spreadsheet, write formulas, format cells, search data, and trace dependencies — all through natural conversation.

Unlike proprietary alternatives, Pi for Excel:
- **Works with any LLM** — Anthropic, OpenAI, Google, local models via Ollama/vLLM
- **Keeps your data local** — the agent runs entirely in the browser; your spreadsheet data never leaves your machine (only the context you send to your chosen LLM provider)
- **Is free and open source** — no subscription, no per-seat pricing

## Why Pi for Excel?

Existing AI add-ins for Excel are closed-source, locked to a single model, and charge $20+/month. They also leave real capabilities on the table:

| | Proprietary add-ins | Pi for Excel |
|---|---|---|
| **Context awareness** | Thin metadata push (sheet names + dimensions). Agent has to make tool calls just to see what you're looking at. | **Rich workbook blueprint** (headers, named ranges, tables, formula density) + **auto-read of your selection** — the agent already knows what you're looking at before you ask. |
| **Formula tracing** | Manual cell-by-cell tracing. Deep dependency trees take dozens of tool calls. | **`trace_dependencies`** — full formula tree in a single call via Office.js `getDirectPrecedents()`. |
| **Sessions** | Total amnesia. Close the sidebar, lose everything. | **Persistent sessions** saved to IndexedDB. Resume any conversation, even after restarting Excel. |
| **Change tracking** | No awareness of what you edited between messages. | **Automatic change tracking** — the agent sees your edits and adapts. |
| **Models** | Locked to one provider and model. | **Any model** — swap between Opus, Sonnet, GPT, Gemini, Codex, or local models mid-conversation. |
| **Cost** | $20+/month per seat. | **Free.** Bring your own API key. |
| **Tool overhead** | Separate tools for compact vs. detailed reads — the model often picks the wrong one. | **11 tools, one per verb.** `read_range` has a `mode` param (compact/csv/detailed). Less overhead, fewer wasted calls. |
| **Writes** | Overwrite protection, but no verification. | **Auto-verification** — reads back written cells to check for `#REF!`, `#VALUE!`, and other errors. |

## Features

- **11 Excel tools** — `get_workbook_overview`, `read_range`, `write_cells`, `fill_formula`, `search_workbook`, `modify_structure`, `format_cells`, `conditional_format`, `comments`, `trace_dependencies`, `view_settings`
- **Composable cell styles** — named format presets (`"currency"`, `"percent"`, `"integer"`) and structural styles (`"header"`, `"total-row"`, `"input"`) that compose like CSS classes: `style: ["currency", "total-row"]`
- **Auto-context injection** — automatically reads around your selection and tracks changes between messages
- **Workbook blueprint** — sends a structural overview of your workbook to the LLM at session start (auto-invalidates after structural changes)
- **Multi-provider auth** — API keys, OAuth (Anthropic, OpenAI, Google, GitHub Copilot, Antigravity), or reuse credentials from Pi TUI
- **Persistent sessions** — conversations auto-save to IndexedDB and survive sidebar close/reopen. Resume any previous session with `/resume`
- **Write verification** — automatically checks formula results after writing
- **Clickable cell references** — cell addresses in assistant messages navigate to the range with a highlight glow
- **Markdown tool cards** — tool outputs render as formatted markdown (tables, lists, headers) instead of raw text
- **Slash commands** — type `/` to browse all available commands with fuzzy search
- **Extensions** — modular extension system with slash commands and inline widget UI (e.g., `/snake`)
- **Keyboard shortcuts** — `Escape` to interrupt, `Shift+Tab` to cycle thinking depth (incl. **max** / `xhigh` effort on Opus 4.6+), `Ctrl+O` to hide/show thinking + tool details
- **Working indicator** — rotating whimsical messages and feature discovery hints while the model is streaming
- **Pi-compatible messages** — conversations use the same `AgentMessage` format as Pi TUI. Session storage differs (IndexedDB vs JSONL), but the message layer is shared — future import/export is straightforward.

## Install (recommended)

See **[`docs/install.md`](./docs/install.md)**.

## Developer Quick Start

### Prerequisites
- Node.js 20+
- Microsoft Excel (desktop, macOS or Windows)
- [mkcert](https://github.com/FiloSottile/mkcert) for local HTTPS

### Setup

```bash
git clone https://github.com/tmustier/pi-for-excel.git
cd pi-for-excel

# Install dependencies
npm install

# Generate HTTPS certificates (required by Office add-ins)
mkcert -install  # one-time: trust the CA
mkcert localhost
mv localhost.pem cert.pem
mv localhost-key.pem key.pem

# Start dev server
npm run dev
```

### Sideload into Excel

**macOS:**
```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

Then open Excel → Insert → My Add-ins → Pi for Excel (Dev).

**Windows:**
```bash
npm run sideload
```

### Configure an LLM provider

On first launch, a welcome overlay appears with provider login options:

1. **OAuth** — click a provider (e.g. Anthropic) to authenticate in your browser, then paste the authorization code back into the add-in when prompted.
2. **API key** — paste a key directly for any supported provider.
3. **Pi TUI credentials** — if you already use [Pi TUI](https://pi.dev), credentials from `~/.pi/agent/auth.json` are loaded automatically in dev mode.

You can change providers later with the `/login` command or by clicking the model name in the status bar.

## Commands

Type `/` in the message input to see all commands:

| Command | Description |
|---------|-------------|
| `/new` | Start a new chat session (current session is saved) |
| `/resume` | Resume a previous session |
| `/name <title>` | Rename the current session |
| `/model` | Switch LLM model |
| `/default-models` | Default model presets (currently opens the model selector) |
| `/login` | Add/change/disconnect API keys and OAuth providers |
| `/settings` | Open settings dialog |
| `/experimental` | Manage experimental feature flags |
| `/extensions` | Open extensions manager (install/enable/disable/reload/uninstall) |
| `/shortcuts` | Show keyboard shortcuts |
| `/compact` | Summarize conversation to free context |
| `/copy` | Copy last response to clipboard |
| `/export` | Export conversation |
| `/share-session` | Share the current session |
| `/snake` | Play Snake! 🐍 (extension) |

Experimental examples:
- `/experimental on tmux-bridge`
- `/experimental tmux-bridge-url https://localhost:3337`
- `/experimental tmux-bridge-token <token>`
- `/experimental tmux-status`
- `/experimental tmux-bridge-url clear`
- `/experimental tmux-bridge-token clear`
- `/experimental on python-bridge`
- `/experimental python-bridge-url https://localhost:3340`
- `/experimental python-bridge-token <token>`
- `/experimental python-bridge-url clear`
- `/experimental python-bridge-token clear`

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Escape` | Interrupt the current response |
| `Shift+Tab` | Cycle thinking depth (off → low → … → max on supported models) |
| `Ctrl+O` | Toggle collapse of thinking blocks and tool messages |
| `/` | Open the slash command menu |

## Architecture

```
src/
├── taskpane.ts              # Thin entrypoint (boot import + tool renderers + bootstrap)
├── boot.ts                  # CSS imports + Lit compat patch install
├── compat/                  # Runtime monkey patches / shims
│   ├── lit-class-field-shadowing.ts
│   └── model-selector-patch.ts
├── taskpane/                # Taskpane wiring modules
│   ├── bootstrap.ts         # Office.onReady + fallback, styles, global patches
│   ├── init.ts              # Agent + sidebar wiring
│   ├── sessions.ts          # IndexedDB session persistence
│   ├── queue-display.ts     # Steering/follow-up queue UI
│   ├── keyboard-shortcuts.ts
│   ├── status-bar.ts
│   ├── welcome-login.ts
│   ├── default-model.ts
│   └── context-injection.ts
├── excel/helpers.ts         # Office.js wrappers + edge-case guards
├── auth/                    # CORS proxy, credential restore, provider mapping
├── tools/                   # Excel tools (read, write, search, format, comments, etc.)
├── conventions/             # Composable cell styles, format presets, style resolver
├── context/                 # Blueprint, selection auto-read, change tracker
├── prompt/system-prompt.ts  # Model-agnostic system prompt builder
├── commands/                # Slash command registry + extensions
│   ├── types.ts             # Command registry + types
│   ├── command-menu.ts      # Slash menu rendering
│   ├── builtins.ts          # Public shim
│   ├── builtins/            # Builtins split by domain (model/settings/session/export/etc.)
│   └── extension-api.ts     # Extension API (overlay, widget, toast, events)
├── extensions/              # Extension modules
│   └── snake.ts             # Snake game (inline widget)
├── ui/                      # Sidebar UI components (Lit + CSS)
│   ├── pi-sidebar.ts
│   ├── pi-input.ts
│   ├── working-indicator.ts
│   ├── tool-renderers.ts    # Render Excel tool output as markdown + collapsible sections
│   ├── theme.css
│   ├── provider-login.ts
│   ├── toast.ts
│   └── loading.ts
└── utils/                   # small shared helpers (content/type guards/errors/format)
```

The agent loop runs client-side in Excel's webview (WebView2 on Windows, WKWebView on Mac). Tool calls execute locally via Office.js — no server round-trips for Excel operations.

## Development

```bash
# Type-check
npx tsc --noEmit

# Build for production
npx vite build

# Validate manifest
npx office-addin-manifest validate manifest.xml
```

### CORS / proxy

**Dev:** the Vite dev server proxies API + OAuth calls to providers (`/api-proxy/*`, `/oauth-proxy/*`).

**Production:** some OAuth/token endpoints are blocked by browser CORS in Office webviews. Pi for Excel supports a **user-configurable CORS proxy**:

1. Start the local proxy (**recommended: HTTPS**):
   ```bash
   npm run proxy:https
   ```
   (defaults to `https://localhost:3001`)

   **Security:** the proxy only accepts browser requests from Pi for Excel origins by default:
   - `https://localhost:3000` (dev)
   - `https://pi-for-excel.vercel.app` (hosted)

   If you host the add-in on a different origin, set `ALLOWED_ORIGINS` (comma-separated):
   ```bash
   ALLOWED_ORIGINS="https://my-addin.example.com" npm run proxy:https
   ```

   By default, the proxy also blocks **loopback + private/local target URLs** to reduce SSRF impact:
   - loopback: `localhost`, `127.0.0.1`, `::1`
   - private/link-local: `10/8`, `172.16/12`, `192.168/16`, `169.254/16`, `fc00::/7`, `fe80::/10`

   Override knobs:
   ```bash
   # Allow loopback targets (legacy local-service behavior)
   ALLOW_LOOPBACK_TARGETS=1 npm run proxy:https

   # Allow private/local targets broadly
   ALLOW_PRIVATE_TARGETS=1 npm run proxy:https

   # Allow only explicit hosts (exact host match, comma-separated)
   ALLOWED_TARGET_HOSTS="api.openai.com,oauth2.googleapis.com" npm run proxy:https

   # Fail closed when DNS lookup fails
   STRICT_TARGET_RESOLUTION=1 npm run proxy:https
   ```

   If port 3001 is taken, pick another port:
   ```bash
   PORT=3003 npm run proxy:https
   ```
2. In Excel, run `/settings` → **Proxy** tab → enable **Use CORS Proxy** → set Proxy URL (e.g. `https://localhost:3003`).

**macOS note:** the taskpane runs on **HTTPS**. WKWebView often blocks calling an **HTTP** proxy from an HTTPS add-in (mixed content). Use `npm run proxy:https` and a proxy URL like `https://localhost:<port>`.

Also: our mkcert cert is for `localhost` by default — `https://127.0.0.1:<port>` will fail unless you generate a cert that includes `127.0.0.1`.

API-key based providers often work without a proxy; OAuth-based logins typically require one.

Proxy rejections return reason codes in plaintext (e.g. `blocked_target_loopback`, `blocked_target_private_ip`, `blocked_target_not_allowlisted`).

### Experimental tmux bridge (local helper)

Start the local bridge:

```bash
# Stub mode (safe default, no real shell execution)
npm run tmux:bridge:https

# Real tmux mode
TMUX_BRIDGE_MODE=tmux npm run tmux:bridge:https
```

Then enable and configure in the add-in:

```bash
/experimental on tmux-bridge
/experimental tmux-bridge-url https://localhost:3337
# optional, if bridge requires bearer auth
/experimental tmux-bridge-token <token>
# diagnostics
/experimental tmux-status
```

Bridge endpoints:
- `GET /health`
- `POST /v1/tmux`

See full request/response contract: [`docs/tmux-bridge-contract.md`](./docs/tmux-bridge-contract.md)

### Experimental Python / LibreOffice bridge (local helper)

Start the local bridge:

```bash
# Stub mode (safe default; deterministic fake responses)
npm run python:bridge:https

# Real local execution mode (runs local python/libreoffice binaries)
PYTHON_BRIDGE_MODE=real npm run python:bridge:https
```

Then enable and configure in the add-in:

```bash
/experimental on python-bridge
/experimental python-bridge-url https://localhost:3340
# optional, if bridge requires bearer auth
/experimental python-bridge-token <token>
```

You can also do this in **/extensions → Local Python / LibreOffice bridge** with one click (`Enable + save URL`).

Bridge endpoints:
- `GET /health`
- `POST /v1/python-run`
- `POST /v1/libreoffice-convert`

Bridge-backed tools:
- `python_run`
- `libreoffice_convert`
- `python_transform_range` (read range → run Python → write output)

For safety, the first Python/LibreOffice bridge execution for a given bridge URL requires user confirmation.

See full request/response contract: [`docs/python-bridge-contract.md`](./docs/python-bridge-contract.md)

### Extension loading safety

`loadExtension()` blocks remote `http(s)` module URLs by default.

- Allowed by default: local module specifiers (`./`, `../`, `/`), blob URLs (used by pasted-code installs), and inline function activators
- Blocked by default: remote extension URLs
- Local specifiers must resolve to bundled extension modules (currently `src/extensions/*.{ts,js}`)
- Temporary unsafe opt-in for local experiments:

```bash
/experimental on remote-extension-urls
```

  (Equivalent low-level toggle: `localStorage.setItem("pi.allowRemoteExtensionUrls", "1")`)

See also: [`docs/extensions.md`](./docs/extensions.md).

## Roadmap

### Shipped in v0.1.0
- [x] 13 Excel tools (read, write, search, format, trace, structure)
- [x] Auto-context injection (workbook blueprint, selection auto-read, change tracking)
- [x] Multi-provider auth (OAuth + API keys for Anthropic, OpenAI, Google, GitHub Copilot, Antigravity)
- [x] Persistent sessions with auto-save and `/resume`
- [x] Write verification (auto-reads back, checks for errors)
- [x] Custom sidebar UI (Lit components, light theme, frosted glass)
- [x] Slash command system with fuzzy search
- [x] Extension system with widget API
- [x] Keyboard shortcuts (Escape, Shift+Tab, Ctrl+O)
- [x] Maintainability refactor — modularized taskpane + slash command builtins

### Shipped in v0.2.0-pre
- [x] Tool consolidation: 14 → 10 tools — one tool per distinct verb, no overlap ([#14](https://github.com/tmustier/pi-for-excel/issues/14) §A, §6)
- [x] `view_settings` tool — gridlines, headings, freeze panes, tab color
- [x] `read_range` gains `mode: "csv"` (absorbs `get_range_as_csv`)
- [x] `get_workbook_overview` gains `sheet` param for sheet-level detail (absorbs `get_all_objects`, closes [#8](https://github.com/tmustier/pi-for-excel/issues/8))
- [x] `search_workbook` gains `context_rows` for surrounding data (closes [#7](https://github.com/tmustier/pi-for-excel/issues/7))
- [x] Compact collapsible tool cards with action verbs + markdown rendering
- [x] Consecutive same-tool grouping with expand/collapse
- [x] Full ESLint upgrade — type-aware `recommendedTypeChecked` preset, 0 errors/warnings
- [x] Architecture: modularized taskpane into 8 focused modules, builtins split by domain

### Shipped in v0.3.0-pre
- [x] Composable cell styles — 6 format presets + 5 structural styles, CSS-like composition ([#1](https://github.com/tmustier/pi-for-excel/issues/1))
- [x] `comments` tool — read, add, update, reply, delete, resolve/reopen cell comments ([#2](https://github.com/tmustier/pi-for-excel/issues/2))
- [x] `read_range` detailed mode surfaces comments within range
- [x] `format_cells` gains `style`, `number_format_dp`, `currency_symbol`, `border_color`, individual border edges
- [x] Markdown rendering in tool output cards ([#15](https://github.com/tmustier/pi-for-excel/issues/15))
- [x] Clickable cell references — navigate to range with highlight glow ([#6](https://github.com/tmustier/pi-for-excel/issues/6) partial)
- [x] Revised welcome copy + expandable hint prompts ([#11](https://github.com/tmustier/pi-for-excel/issues/11))
- [x] Humanized tool card inputs/outputs (color names, format labels)
- [x] Blueprint invalidation after structural changes
- [x] UI polish: queue layout, thinking/tool card styling, case-insensitive model search
- [x] Experimental tmux bridge: `/experimental` feature flags + local helper + gated `tmux` tool ([#3](https://github.com/tmustier/pi-for-excel/issues/3))
- [x] Experimental Python / LibreOffice bridge: gated `python_run` + `libreoffice_convert` + `python_transform_range`, local helper, first-run approval per bridge URL, and `/experimental python-bridge-*` config ([#25](https://github.com/tmustier/pi-for-excel/issues/25))

### Up next
- [ ] New tools: charts, tables, data validation ([#18](https://github.com/tmustier/pi-for-excel/issues/18))
- [ ] Progressive disclosure ([#18](https://github.com/tmustier/pi-for-excel/issues/18)) — on-demand tool injection as tool count grows
- [ ] Conventions Phase 2 ([#1](https://github.com/tmustier/pi-for-excel/issues/1)) — user-configurable via settings UI, workbook-scoped
- [ ] Native Excel styles vs. custom style system ([#19](https://github.com/tmustier/pi-for-excel/issues/19))
- [ ] Auto-compaction ([#20](https://github.com/tmustier/pi-for-excel/issues/20)) — context window budget management for long conversations
- [ ] Change approval UI ([#6](https://github.com/tmustier/pi-for-excel/issues/6)) — structured approval flow for overwrites
- [ ] Header bar UX ([#12](https://github.com/tmustier/pi-for-excel/issues/12)) — session switcher, workbook indicator
- [ ] Extension platform follow-ups ([#13](https://github.com/tmustier/pi-for-excel/issues/13)) — sandbox/permissions, widget API evolution, docs polish

### Future
- [ ] Production CORS solution ([#4](https://github.com/tmustier/pi-for-excel/issues/4)) — service worker or hosted relay
- [ ] Distribution: non-technical install ([#16](https://github.com/tmustier/pi-for-excel/issues/16)) — hosted build + production manifest
- [ ] Python code execution via Pyodide
- [ ] SpreadsheetBench evaluation (target >43%)
- [ ] Per-workbook instructions (like AGENTS.md)
- [ ] On-demand tier 2 tools: named ranges, protection, page layout, images, hyperlinks
- [ ] Pivot tables and slicers
- [ ] Pi TUI ↔ Excel session import/export

## Prior Art & Credits

- [Pi](https://pi.dev) by [@badlogic](https://github.com/badlogic) (Mario Zechner) — the agent framework powering this project (source: https://github.com/badlogic/pi-mono). Pi for Excel uses pi-agent-core, pi-ai, and pi-web-ui for the agent loop, LLM abstraction, and session storage.
- [whimsical.ts](https://github.com/mitsuhiko/agent-stuff/blob/main/pi-extensions/whimsical.ts) by [@mitsuhiko](https://github.com/mitsuhiko) (Armin Ronacher) — the rotating "Working…" messages are adapted from his Pi extension, rewritten for a spreadsheet/finance audience.
- [Microsoft Copilot Agent Mode](https://techcommunity.microsoft.com/) — JS code gen + reflection, 57.2% SpreadsheetBench
- [Univer](https://univer.ai) — Canvas-based spreadsheet runtime, 68.86% SpreadsheetBench (different architecture)

## License

MIT
