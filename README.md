# Pi for Excel

An open-source, multi-model AI sidebar add-in for Microsoft Excel — powered by [Pi](https://github.com/mariozechner/pi-coding-agent).

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
| **Tool overhead** | Separate tools for compact vs. detailed reads — the model often picks the wrong one. | **Single `read_range` tool** with a `mode` parameter. Less overhead, fewer wasted calls. |
| **Writes** | Overwrite protection, but no verification. | **Auto-verification** — reads back written cells to check for `#REF!`, `#VALUE!`, and other errors. |

## Features

- **13 Excel tools** — `get_workbook_overview`, `read_range`, `get_range_as_csv`, `read_selection`, `get_all_objects`, `write_cells`, `fill_formula`, `search_workbook`, `modify_structure`, `format_cells`, `conditional_format`, `trace_dependencies`, `get_recent_changes`
- **Auto-context injection** — automatically reads around your selection and tracks changes between messages
- **Workbook blueprint** — sends a structural overview of your workbook to the LLM at session start
- **Multi-provider auth** — API keys, OAuth (Anthropic, OpenAI, Google, GitHub Copilot, Antigravity), or reuse credentials from Pi TUI
- **Persistent sessions** — conversations auto-save to IndexedDB and survive sidebar close/reopen. Resume any previous session with `/resume`
- **Write verification** — automatically checks formula results after writing
- **Slash commands** — type `/` to browse all available commands with fuzzy search
- **Extensions** — modular extension system with slash commands and inline widget UI (e.g., `/snake`)
- **Keyboard shortcuts** — `Escape` to interrupt, `Shift+Tab` to focus input, `Ctrl+O` to collapse thinking/tool blocks
- **Working indicator** — rotating whimsical messages and feature discovery hints while the model is streaming
- **Pi TUI interop** — sessions use the same `SessionData` format as pi-web-ui — future export/import is free

## Quick Start

### Prerequisites
- Node.js 18+
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
npx vite --port 3000
```

### Sideload into Excel

**macOS:**
```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```

Then open Excel → Insert → My Add-ins → Pi for Excel (Dev).

**Windows:**
```bash
npx office-addin-debugging start manifest.xml desktop --app excel
```

### Configure an LLM provider

On first launch, a welcome overlay appears with provider login options:

1. **OAuth** — click a provider (Anthropic, Google) to authenticate in your browser.
2. **API key** — paste a key directly for any supported provider.
3. **Pi TUI credentials** — if you already use [Pi TUI](https://github.com/mariozechner/pi-coding-agent), credentials from `~/.pi/agent/auth.json` are loaded automatically in dev mode.

You can change providers later with the `/login` command or by clicking the model name in the status bar.

## Commands

Type `/` in the message input to see all commands:

| Command | Description |
|---------|-------------|
| `/new` | Start a new chat session (current session is saved) |
| `/resume` | Resume a previous session |
| `/name <title>` | Rename the current session |
| `/model` | Switch LLM model |
| `/default-models` | Set preferred models per provider |
| `/login` | Add or change API keys / OAuth |
| `/settings` | Open settings dialog |
| `/shortcuts` | Show keyboard shortcuts |
| `/compact` | Summarize conversation to free context |
| `/copy` | Copy last response to clipboard |
| `/export` | Export conversation |
| `/share-session` | Share the current session |
| `/snake` | Play Snake! 🐍 (extension) |

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Escape` | Interrupt the current response |
| `Shift+Tab` | Focus the input field |
| `Ctrl+O` | Toggle collapse of thinking blocks and tool messages |
| `/` | Open the slash command menu |

## Architecture

```
src/
├── taskpane.ts            # Entry — mounts sidebar, wires agent, status bar
├── boot.ts                # Lit class field fix + CSS imports
├── excel/helpers.ts       # Office.js wrappers + edge-case guards
├── auth/                  # CORS proxy, credential restore, provider mapping
├── tools/                 # 13 Excel tools (read, write, search, format, etc.)
├── context/               # Blueprint, selection auto-read, change tracker
├── prompt/system-prompt.ts # Model-agnostic system prompt builder
├── commands/              # Slash command registry, builtins, extension API
│   ├── types.ts           # Command registry + types
│   ├── builtins.ts        # Built-in slash commands
│   ├── command-menu.ts    # Slash menu rendering
│   └── extension-api.ts   # Extension API (overlay, widget, toast, events)
├── extensions/            # Extension modules
│   └── snake.ts           # Snake game (inline widget)
├── ui/                    # Sidebar UI components (Lit + CSS)
│   ├── pi-sidebar.ts      # Main layout (messages, input, widget slot)
│   ├── pi-input.ts        # Chat input with auto-grow + placeholder rotation
│   ├── working-indicator.ts # Streaming status with rotating messages
│   ├── theme.css          # Light theme, glass effects, component styles
│   ├── provider-login.ts  # OAuth + API key login rows
│   ├── toast.ts           # Toast notifications
│   └── loading.ts         # Loading spinner + error banner
└── utils/format.ts        # Markdown tables, token truncation
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

### CORS in development

The Vite dev server proxies API calls to LLM providers, stripping browser headers that would trigger CORS failures (notably Anthropic rejects requests with `Origin` headers). This is dev-only — production deployment will need a different solution.

## Roadmap

- [ ] Python code execution via Pyodide
- [ ] SpreadsheetBench evaluation (target >43%)
- [ ] Production CORS solution (service worker or hosted relay)
- [ ] Per-workbook instructions (like AGENTS.md)
- [ ] Chart creation and modification
- [ ] Named range awareness in formulas
- [ ] Data validation
- [ ] Pi TUI ↔ Excel session teleport
- [ ] Extension API build-out (#13) — dynamic loading, tool registration, sandboxing
- [ ] Header bar UX (#12) — session switcher, workbook indicator

## Prior Art

- [Microsoft Copilot Agent Mode](https://techcommunity.microsoft.com/) — JS code gen + reflection, 57.2% SpreadsheetBench
- [Univer](https://univer.ai) — Canvas-based spreadsheet runtime, 68.86% SpreadsheetBench (different architecture)

## License

MIT
