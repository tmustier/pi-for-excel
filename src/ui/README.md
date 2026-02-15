# UI Architecture

## Layout

The sidebar UI has two layers:

1. **Our components** (`pi-sidebar.ts`, `pi-input.ts`) — own the layout shell (scroll area + input footer). Purpose-built for ~350px.
2. **pi-web-ui content components** — render message internals (markdown, code blocks, tool cards, thinking blocks). Registered via `src/ui/register-components.ts` (deep imports from `@mariozechner/pi-web-ui/dist/*`).

```
┌─ pi-sidebar ──────────────────────────────────────┐
│  .pi-messages          ← scrollable               │
│    message-list        ← pi-web-ui                │
│    streaming-message-container  ← pi-web-ui       │
│    .pi-empty           ← empty state overlay       │
│  .pi-working           ← "Working…" pulse (stream) │
│  .pi-input-area        ← sticky footer            │
│    pi-input            ← our component             │
│    #pi-status-bar      ← model + ctx % + thinking  │
└────────────────────────────────────────────────────┘
```

`pi-sidebar` subscribes to the `Agent` directly and passes messages/tools/streaming state down as properties to the pi-web-ui components.

## Styling

**Two CSS files, loaded in order** (see `boot.ts`):

1. `@mariozechner/pi-web-ui/app.css` — Tailwind v4 (utilities in `@layer`)
2. `./ui/theme.css` — our variables, component styles, and content overrides

### The critical rule

> **Never add unlayered `margin: 0` or `padding: 0` to a universal selector.**

Tailwind v4 puts all utilities inside `@layer utilities`. Unlayered CSS always beats layered CSS regardless of specificity. A bare `* { padding: 0 }` silently zeros out every `py-2`, `px-4`, `p-2.5` etc. in pi-web-ui. The `taskpane.html` inline `<style>` intentionally only sets `box-sizing: border-box` on `*`.

### Theme guardrails

- `npm run check:css-theme` verifies every `var(--token)` used in local theme CSS resolves to a defined custom property (or has an explicit fallback).
- `npm run check:theme-utility-overrides` blocks Tailwind utility-class selectors in theme modules (except `theme/unstable-overrides.css`).
- `npm run check:builtins-inline-style` blocks inline `style.*` usage in `src/commands/builtins/**` so overlay styling stays class-based.

### theme.css structure

| Section | What it does |
|---|---|
| 1. CSS Variables | Colors, fonts, glass tokens — pi-web-ui consumes these via `var(--background)` etc. |
| 2. Global | Body background (spreadsheet grid texture), scrollbars |
| 3–5. Our components | `.pi-messages`, `.pi-input-card`, `.pi-empty` — fully ours, no overrides needed |
| 6. Working indicator | `.pi-working` — pulsing "Working…" bar shown during streaming |
| 7–10. Chrome | Status bar (model picker + ctx + thinking), toast, slash command menu, welcome overlay |
| 10b. Overlay primitives | Shared classes for builtins overlays (tabs, textarea, buttons, footer actions) |
| 11. Content overrides | **Targeted** pi-web-ui tweaks — user bubble color, sidebar-width margins, tool card borders, semantic classes from message style hooks |
| 12–13. Dialogs, unstable overrides, Queue | Stable dialog styling via runtime hooks + (currently empty) unstable override buffer + steer/follow-up queue |

> Note: `theme.css` is an entrypoint; styles are split into `src/ui/theme/*.css` and imported in order:
> - `theme/tokens.css` (1)
> - `theme/base.css` (2)
> - `theme/components.css` (3–10, import-only entrypoint) → imports `theme/components/{tabs,input,empty-state,working-indicator,widgets,status-bar,toasts,menus,welcome,files,welcome-login}.css`
> - `theme/overlays.css` (10b) → imports `theme/overlays/{primitives,extensions,integrations,skills,provider-resume-shortcuts,recovery,experimental}.css`
> - `theme/content-overrides.css` (11) → imports `theme/content/{messages,tool-cards,csv-table,dependency-tree,tool-card-markdown,message-components}.css`
> - `theme/dialogs.css` (12, stable selectors)
> - `theme/unstable-overrides.css` (12b, utility-coupled upstream selectors)
> - `theme/queue.css` (13)

### When overriding pi-web-ui styles

pi-web-ui uses Light DOM (`createRenderRoot() { return this; }`), so styles leak both ways. When you need to override:

- **Prefer CSS variables** (`--background`, `--border`, `--primary`, etc.) — pi-web-ui reads these.
- **Use element-scoped selectors** like `user-message .mx-4` or `tool-message .border` — not bare class names.
- For message internals, prefer adding semantic classes in `src/ui/message-style-hooks.ts` (`applyMessageStyleHooks`) and target those classes in CSS.
- **Use `!important` sparingly** — only needed when overriding Tailwind utility classes that also use `!important` or when specificity within `@layer` can't be beaten otherwise.
- **Don't target deep Tailwind internals** like `.px-2.pb-2 > .flex.gap-2:last-child > button:last-child`. These break on library updates. Target the custom element tag or a stable class name.
- If you must target utility internals, place the rule in `src/ui/theme/unstable-overrides.css` with a short comment.

## Components

| File | Replaces | Notes |
|---|---|---|
| `pi-sidebar.ts` | ChatPanel + AgentInterface | Owns layout, subscribes to Agent, renders message-list + streaming container + working indicator |
| `pi-input.ts` | MessageEditor | Auto-growing textarea, send/abort buttons, `+` input actions menu, file import affordances; fires `pi-send` / `pi-abort` / `pi-files-drop` / `pi-input-action` events |
| `message-style-hooks.ts` | — | Stamps semantic classes on pi-web-ui message internals (`pi-assistant-body`, `pi-tool-card-fallback`, etc.) to avoid brittle utility selectors |
| `dialog-style-hooks.ts` | — | Stamps semantic classes on dialog internals (`pi-dialog-card`, `pi-model-selector-item-*`) so dialog CSS avoids utility selectors |
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
