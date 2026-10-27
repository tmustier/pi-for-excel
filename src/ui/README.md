# UI Architecture

## Layout

The sidebar UI is first-party end to end:

1. **Shell components** (`pi-sidebar.ts`, `pi-input.ts`) — own the layout shell (scroll area + input footer). Purpose-built for ~350px.
2. **Message components** (`src/ui/messages/`) — render message internals (markdown, code blocks, tool cards, thinking blocks, attachments). Registered via `src/ui/register-components.ts`. Clean-roomed from pi-web-ui 0.75.3 during the UI ownership migration (`docs/ui-ownership.md`).

```
┌─ pi-sidebar ──────────────────────────────────────┐
│  .pi-session-tabs      ← tab strip (scrolls)      │
│  #pi-notices           ← banners (banner.ts)      │
│  .pi-messages          ← scrollable               │
│    message-list        ← src/ui/messages          │
│    streaming-message-container  ← src/ui/messages │
│    .pi-empty           ← empty state              │
│  .pi-dock              ← composer unit            │
│    .pi-working         ← working line (stream)    │
│    queue / widgets     ← steer + follow-up tray   │
│    .pi-input-area      ← pi-input                 │
│    #pi-status-bar      ← model, thinking, ctx,    │
│                          mode                     │
└────────────────────────────────────────────────────┘
```

`pi-sidebar` subscribes to the `Agent` directly and passes messages/tools/streaming state down as properties to the message components.

One model of a streaming turn: the response being streamed renders only in
`<streaming-message-container>`. At `message_end` pi-agent-core pushes the
finished response into `state.messages`; the sidebar clears the streaming copy
in the same update, so `<message-list>` alone draws it from then on. Its tool
cards spin until their own `tool_execution_end` (results arriving before the
batch's toolResult messages are passed down as `liveToolResults`). Whether a
card, thinking block or group is open comes from `transcript-state.ts`, never
from the element, so it survives the element being replaced.

### Message component modules (`src/ui/messages/`)

| Module | Elements / exports |
|---|---|
| `message-list.ts` | `<message-list>` — stable history, dispatches to custom role renderers first |
| `streaming-message-container.ts` | `<streaming-message-container>` — rAF-batched streaming message (`setMessage()`) |
| `messages.ts` | `<user-message>`, `<assistant-message>`, `<tool-message>` |
| `markdown-block.ts` | `<markdown-block>` — marked + input hardening; fenced code → `<code-block encoding="base64">` |
| `code-block.ts` | `<code-block>` — highlight.js core + copy button |
| `thinking-block.ts` | `<thinking-block>` — collapsible (header is a button); "Thinking…" while it is the streaming block, then "Thought for Xs" from transcript state, "Thought" when the timing was not observed |
| `transcript-state.ts` | View state that outlives elements: details mode (Ctrl/Cmd+O, saved), hand toggles of cards/blocks/groups, observed thinking durations |
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
| 12–14. Dialogs, Queue, Legacy WebView | First-party dialog styling (model selector, provider connect) + steer/follow-up queue + sRGB fallbacks for WPS/older WebViews without OKLCH support |

> Note: `theme.css` is an entrypoint; styles are split into `src/ui/theme/*.css` and imported in order:
> - `theme/tokens.css` (1)
> - `theme/preflight.css` (1b)
> - `theme/base.css` (2)
> - `theme/components.css` (3–10, import-only entrypoint) → imports `theme/components/{tabs,input,empty-state,working-indicator,widgets,status-bar,toasts,menus,welcome,files,welcome-login}.css`
> - `theme/overlays.css` (10b) → imports `theme/overlays/{primitives,extensions,integrations,skills,provider-resume-shortcuts,recovery,experimental}.css`
> - `theme/content-overrides.css` (11) → imports `theme/content/{messages,tool-cards,csv-table,dependency-tree,tool-card-markdown,message-components}.css`
> - `theme/dialogs.css` (12, stable selectors)
> - `theme/queue.css` (13)
> - `theme/legacy-webview.css` (14, sRGB fallbacks under `@supports not (color: oklch(...))`)

### Radius system

All radii come from the token scale in `theme/tokens.css` — never hardcode.
Tiers, by surface role:

| Token | Value | Used for |
|---|---|---|
| `--radius-xl` | 20px | Overlay dialogs (outermost surfaces) |
| `--pill-radius` (= `--radius-lg`) | 16px | Transcript units: standalone tool cards, tool groups, user bubble, input card |
| `--radius-md` | 12px | Cards *inside* another surface (list cards in overlays, rows inside 16px pills) |
| `--radius-sm` | 8px | Buttons, command blocks, small interactive chrome |
| `--radius-xs` | 4px | Chips, badges, tiny inline elements |
| `--radius-full` | round | Circular buttons, pill badges, toggle knobs |

**Concentric rule for nesting:** when an element with a visible
background/border sits inside a rounded parent, its radius should be
*parent radius − inset*, clamped to the nearest token. Examples:

- Grouped tool-card rows sit 4px inside a 16px group → 12px (`--radius-md`).
- List cards sit ~8px inside a 20px dialog → 12px (`--radius-md`).
- Matching the parent's radius on a nested element (16 inside 16) reads
  wrong at the corners — avoid it.

Standalone vs grouped tool cards intentionally share the same outer
language: **one 16px pill per transcript unit** (a lone card or a whole
group), with grouped rows demoted to inset 12px rows inside the pill.

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
| `toast.ts` | — | `showToast(msg, duration \| { duration, variant })` + `showActionToast({ id?, … })` + `dismissToast(id)` — stacked transient notifications (see Notifications) |
| `banner.ts` | — | `showBanner` / `createBanner` / `mountBanner` / `hideBanner` — persistent conditions in `#pi-notices` (see Notifications) |
| `announcer.ts` | — | `announce(msg, { assertive? })` — hidden live regions (see Notifications) |
| `menu-keyboard.ts` | — | Shared `role="menu"` keyboard model (see Menus) |
| `text-input-dialog.ts` | `window.prompt` | `requestTextInputDialog(...)` on the overlay primitive; native `prompt`/`confirm`/`alert` are dead on Excel on the web |
| `theme-mode.ts` | — | Keeps light mode by default; `/experimental on dark-mode` enables Office/theme-driven `.dark` (fallback: `prefers-color-scheme`) |
| `loading.ts` | — | Splash screen shown during init |
| `provider-login.ts` | — | API key entry rows for the welcome overlay |
| `overlay-dialog.ts` | — | Shared overlay lifecycle helper (single-instance toggle, Escape/backdrop close, focus contract — see Overlays) + shared overlay chrome/DOM builders (`createOverlayCloseButton`, `createOverlayHeader`, `createOverlayButton`, `createOverlayInput`, `createOverlaySectionTitle`) |
| `overlay-ids.ts` | — | Shared overlay id constants used across builtins/extensions/taskpane |

## Layers

One z-index scale in `theme/tokens.css`; never use a raw number:

| Token | Used by |
|---|---|
| `--z-sticky` | tab strip, dock, status bar |
| `--z-popover` | menus, status popovers, command menu |
| `--z-overlay` | `.pi-overlay` (every dialog and full-pane view) |
| `--z-toast` | `#pi-toast` — above overlays, so feedback from inside a dialog is visible |

## Overlays

Every dialog and full-pane view (Settings, Files, Plugins, resume, model picker, confirmations, text input, extension overlays) is built with `createOverlayDialog` and gets the same contract from `installOverlayFocus`:

- **Initial focus**: `[data-autofocus]`, else the first text field, else the card itself (so the dialog is announced and nothing is activated by accident).
- **Focus trap**: Tab / Shift+Tab cycle inside the topmost overlay; focus that lands behind it is pulled back. Toasts stay reachable.
- **Escape** acts on the topmost overlay only. **Back/close rule**: Escape returns to where the user came from (Files preview → list, a Settings page reached from the Settings root → root) and closes the overlay on the view it opened at (`/rules`, `/backups`). A view's `onEscape` returns `true` when it handled the key. A task's own Save/Cancel finish the task and close. Settings' ‹ button is "up" and may go to a page's `parentId` on the entry page. × and backdrop click close through the same guarded path, so unsaved-change confirmations apply.
- **Return focus** to the opener, or the composer when the opener is gone.
- While any modal overlay is open, taskpane shortcuts (`keyboard-shortcuts.ts`) do nothing; older overlays are made inert (`coverOpenOverlays`).

Do not call `window.prompt` / `confirm` / `alert`: use `requestTextInputDialog` / `requestConfirmationDialog`.

## Notifications

Pick by how long the thing stays true:

| Kind | Use for | Where | Rules |
|---|---|---|---|
| **Inline status** | state of one control or card (saving…, tool failed) | next to the control | no extra notification for the same event |
| **Toast** (`toast.ts`) | the result of something the user just did | `#pi-toast`, bottom, above the dock and above overlays | ≤3 stacked, newest last; duplicates refresh instead of stacking; pause on hover/focus; errors stay 6 s and have ×; Undo toasts take an `id` so they can be dismissed when the action is undone another way |
| **Banner** (`banner.ts`) | a condition that stays true until fixed or acknowledged (proxy down, first-run disclosure) | `#pi-notices` under the tabs, most severe first | always dismissible; one per `id`; never moves focus; `role="alert"` only for errors |
| **Announcement** (`announcer.ts`) | screen-reader-only news with no visible notification: a run finished, failed in a background tab, is retrying; an approval is waiting | hidden `aria-live` regions | never for something a toast or banner already announces; assertive only when the user must act |

Wording: sentence case, past tense for results ("Renamed tab to “Forecast”"), say what to do next for failures ("Couldn’t save. Try again."), curly quotes around names, no trailing period on one-phrase toasts, and the same string for the same event on every path. Every string goes through `t()`.

## Menus

`menu-keyboard.ts` is the one keyboard model for `role="menu"` (gear menu, tab context menu, status popovers):

- Opening focuses the checked item (radio menus) or the first item.
- ↑/↓ move with wrap, Home/End jump, Enter/Space activate, Escape and Tab close; closing returns focus to the trigger.
- Tab context menu also opens with Shift+F10 / the ContextMenu key.
- The slash-command menu is a combobox on the composer (`aria-activedescendant`); it only takes keys typed in the composer.

## Wiring (taskpane.ts)

`taskpane.ts` creates the `Agent`, mounts `PiSidebar`, and wires:
- `sidebar.onSend` / `sidebar.onAbort` → agent.prompt() / agent.abort()
- Keyboard shortcuts (Enter to steer, Alt/Option+Enter for follow-up, Escape to stop, Ctrl+Alt/Option+T for thinking) via `document.addEventListener("keydown")`; Shift+Tab keeps native focus navigation.
- Slash command menu via `wireCommandMenu(sidebar.getInput())`
- Session persistence (auto-save on message_end, auto-restore latest on init)
