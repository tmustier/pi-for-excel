# UI Polish — Frosted Glass Design

Pivoted from full rewrite to CSS-driven design on top of pi-web-ui.
Keep pi-web-ui for chat functionality (streaming, auto-scroll, sessions).
Override styling comprehensively via CSS. Build our own empty state + header.

## Architecture Decision
- **Keep** pi-web-ui ChatPanel for chat logic (streaming, tool rendering, message list)
- **Override** all visual styling via `theme.css` (light DOM = full CSS control)
- **Replace** model picker via `onModelSelect` callback (future)
- **Build** our own: header, empty state, loading states

## Done
- [x] Frosted glass design (backdrop-filter blur on header, input, dialogs, messages)
- [x] Spreadsheet grid background pattern (teal cell lines)
- [x] Empty state with π logo, tagline, and hint cards
- [x] Input card: translucent glass with subtle focus ring
- [x] Model picker: frosted glass dialog, compact layout, sans-serif fonts
- [x] User messages: frosted tinted pills with inner glow
- [x] Tool cards: translucent with blur
- [x] Header: frosted glass with teal accent line
- [x] Comprehensive typography overrides (DM Sans + JetBrains Mono)
- [x] All dialogs: frosted glass treatment

## Remaining Polish
- [x] Make hint cards clickable (populate input with suggestion)
- [x] Polish tuning pass (spacing, opacity, hover/active states)
- [x] Manifest copied to sideload path
- [ ] Sideload in Excel and verify glass effect with real spreadsheet behind
- [ ] Consider our own model picker component (intercept onModelSelect)
- [ ] Awaiting user feedback from live Excel testing
