# UI Architecture

## Layout

The sidebar UI is first-party end to end:

1. **Shell components** (`pi-sidebar.ts`, `pi-input.ts`) — own the layout shell (scroll area + input footer). Purpose-built for ~350px.
2. **Message components** (`src/ui/messages/`) — render message internals (markdown, code blocks, tool cards, thinking blocks, attachments). Registered via `src/ui/register-components.ts`. Clean-roomed from pi-web-ui 0.75.3 during the UI ownership migration (`docs/ui-ownership.md`).

```
┌─ pi-sidebar ──────────────────────────────────────┐
│  .pi-messages          ← scrollable               │
│    message-list        ← src/ui/messages                │
│    streaming-message-container  ← src/ui/messages       │
│    .pi-empty           ← empty state overlay       │
│  .pi-working           ← "Working…" pulse (stream) │
│  .pi-input-area        ← sticky footer            │
│    pi-input            ← our component             │
│    #pi-status-bar      ← model + ctx % + thinking  │
└────────────────────────────────────────────────────┘
```

`pi-sidebar` subscribes to the `Agent` directly and passes messages/tools/streaming state down as properties to the message components.

### Message component modules (`src/ui/messages/`)

| Module | Elements / exports |
|---|---|
| `message-list.ts` | `<message-list>` — stable history, dispatches to custom role renderers first |
| `streaming-message-container.ts` | `<streaming-message-container>` — rAF-batched streaming message (`setMessage()`) |
| `messages.ts` | `<user-message>`, `<assistant-message>`, `<tool-message>` |
| `markdown-block.ts` | `<markdown-block>` — marked + input hardening; fenced code → `<code-block encoding="base64">` |
| `code-block.ts` | `<code-block>` — highlight.js core + copy button |
| `thinking-block.ts` | `<thinking-block>` — collapsible; owns "Thinking…" → "Thought for Xs" label lifecycle |
| `attachment-tile.ts` | `<attachment-tile>` — defensive rendering for restored attachment messages |
| `message-renderer-registry.ts` | `registerMessageRenderer()` / `renderMessage()` by role |
| `tool-renderer-registry.ts` | `registerToolRenderer()` / `renderTool()` + JSON fallback renderer |

Security invariants for `<markdown-block>` live in `src/compat/marked-safety.ts`
(prototype-level link/image hardening) — keep both in sync.

## Styling

**One CSS entrypoint** (see `boot.ts`): `./ui/theme.css` — tokens, first-party
preflight reset, component styles, and content styles. Tailwind and
`pi-web-ui/app.css` are gone; there is no layered CSS left.

### Preflight

`theme/preflight.css` replaces Tailwind v4's preflight. The component CSS was
written against those reset semantics (universal `margin`/`padding` zeroing,
border-style reset, `display: block` media, form-control font inheritance), so
keep it faithful — removing rules from it can un-hide subtle spacing
regressions across every component. It also owns the `.hidden` / `[hidden]`
visibility utilities that shell components toggle.

### Theme guardrails

- `npm run check:css-theme` verifies every `var(--token)` used in local theme CSS resolves to a defined custom property (or has an explicit fallback).
- `npm run check:theme-utility-overrides` blocks Tailwind-style utility-class selectors in all theme modules (templates must use semantic classes).
- `npm run check:builtins-inline-style` blocks inline `style.*` usage in `src/commands/builtins/**` so overlay styling stays class-based.

### theme.css structure

| Section | What it does |
|---|---|
| 1. CSS Variables | Colors, fonts, glass tokens — consumed everywhere via `var(--background)` etc. |
| 1b. Preflight | First-party base reset (`theme/preflight.css`) |
| 2. Global | Body background (spreadsheet grid texture), scrollbars |
| 3–5. Our components | `.pi-messages`, `.pi-input-card`, `.pi-empty` — fully ours, no overrides needed |
| 6. Working indicator | `.pi-working` — pulsing "Working…" bar shown during streaming |
| 7–10. Chrome | Status bar (model picker + ctx + thinking), toast, slash command menu, welcome overlay |
| 10b. Overlay primitives | Shared classes for builtins overlays (tabs, textarea, buttons, footer actions) |
| 11. Content styles | Message component styling — user bubble, sidebar-width margins, tool cards, markdown/code/thinking styles (all first-party semantic classes) |
| 12–13. Dialogs, Queue | First-party dialog styling (model selector, provider connect) + steer/follow-up queue |

> Note: `theme.css` is an entrypoint; styles are split into `src/ui/theme/*.css` and imported in order:
> - `theme/tokens.css` (1)
> - `theme/preflight.css` (1b)
> - `theme/base.css` (2)
> - `theme/components.css` (3–10, import-only entrypoint) → imports `theme/components/{tabs,input,empty-state,working-indicator,widgets,status-bar,toasts,menus,welcome,files,welcome-login}.css`
> - `theme/overlays.css` (10b) → imports `theme/overlays/{primitives,extensions,integrations,skills,provider-resume-shortcuts,recovery,experimental}.css`
> - `theme/content-overrides.css` (11) → imports `theme/content/{messages,tool-cards,csv-table,dependency-tree,tool-card-markdown,message-components}.css`
> - `theme/dialogs.css` (12, stable selectors)
> - `theme/queue.css` (13)

### Styling message components

Message components use Light DOM (`createRenderRoot() { return this; }`), so theme CSS applies directly. Conventions:

- Templates carry **semantic classes** (`pi-assistant-body`, `pi-code__bar`, `pi-thinking-label`, …) — never Tailwind utilities.
- Style them in `src/ui/theme/content/*.css` using design tokens from `theme/tokens.css`.
- Custom-element tags (`markdown-block`, `thinking-block`, …) are stable selectors; prefer them for typography scoping.

## Components

| File | Replaces | Notes |
|---|---|---|
| `pi-sidebar.ts` | ChatPanel + AgentInterface | Owns layout, subscribes to Agent, renders message-list + streaming container + working indicator |
| `pi-input.ts` | MessageEditor | Auto-growing textarea, send/abort buttons, `+` input actions menu, file import affordances; fires `pi-send` / `pi-abort` / `pi-files-drop` / `pi-input-action` events |
| `messages/*` | pi-web-ui Messages, MessageList, StreamingMessageContainer, ThinkingBlock, AttachmentTile + mini-lit MarkdownBlock/CodeBlock | First-party message rendering (see table above) |
| `icons.ts` | mini-lit `icon`/`iconDOM` | Lucide icon helpers sized via SVG attributes |
| `model-selector-dialog.ts` | pi-web-ui ModelSelector | First-party model picker (active-provider filter, featured ordering, search, keyboard nav) |
| `api-key-dialog.ts` | pi-web-ui ApiKeyPromptDialog | First-party provider connect prompt (reuses `provider-login.ts` rows, so OAuth works too) |
| `toast.ts` | — | `showToast(msg, duration \| { duration, variant })` + `showActionToast(...)` — fixed notifications with destructive styling for errors |
| `theme-mode.ts` | — | Keeps light mode by default; `/experimental on dark-mode` enables Office/theme-driven `.dark` (fallback: `prefers-color-scheme`) |
| `loading.ts` | — | Splash screen shown during init |
| `provider-login.ts` | — | API key entry rows for the welcome overlay |
| `overlay-dialog.ts` | — | Shared overlay lifecycle helper (single-instance toggle, Escape/backdrop close, focus restore) + shared overlay chrome/DOM builders (`createOverlayCloseButton`, `createOverlayHeader`, `createOverlayButton`, `createOverlayInput`, `createOverlaySectionTitle`, `createOverlayBadge`) |
| `overlay-ids.ts` | — | Shared overlay id constants used across builtins/extensions/taskpane |

## Wiring (taskpane.ts)

`taskpane.ts` creates the `Agent`, mounts `PiSidebar`, and wires:
- `sidebar.onSend` / `sidebar.onAbort` → agent.prompt() / agent.abort()
- Keyboard shortcuts (Enter, Escape, Shift+Tab for thinking) via `document.addEventListener("keydown")`
- Slash command menu via `wireCommandMenu(sidebar.getTextarea())`
- Session persistence (auto-save on message_end, auto-restore latest on init)
