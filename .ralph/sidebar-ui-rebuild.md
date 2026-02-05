# Replace pi-web-ui layout shell with custom sidebar components

## Context
Pi for Excel uses pi-web-ui's ChatPanel/AgentInterface/MessageEditor as the layout shell. These are designed for full-width desktop apps, not a ~350px Excel sidebar. The current approach fights the library with 550 lines of fragile CSS overrides. 

## Approach
Replace the layout shell (ChatPanel, AgentInterface, MessageEditor) with purpose-built Lit components, while **keeping** pi-web-ui's content rendering components (user-message, assistant-message, tool-message, message-list, streaming-message-container, markdown-block, thinking-block, code-block, ModelSelector, ApiKeyPromptDialog).

## Design Direction
"Frosted Sidebar" — refined minimalism with depth. Translucent surfaces over a subtle spreadsheet grid texture. DM Sans + JetBrains Mono. Teal-green accent. Glass effects give depth without being heavy. Every element purpose-built for ~350px width.

## Checklist
- [x] Create `src/ui/pi-sidebar.ts` — Main sidebar Lit component (replaces ChatPanel + AgentInterface)
- [x] Create `src/ui/pi-input.ts` — Input component (replaces MessageEditor)  
- [x] Rewrite `src/ui/theme.css` — Clean CSS: variables + our component styles + minimal content component overrides
- [x] Update `src/ui/index.ts` — New exports
- [x] Update `src/boot.ts` — No changes needed (already imports app.css then theme.css correctly)
- [x] Rewrite `src/taskpane.ts` — Use PiSidebar instead of ChatPanel, wire keyboard shortcuts/commands/sessions
- [x] Update `src/dialog.ts` — Decision: keep ChatPanel for pop-out (full-width window where default layout works)
- [x] Verify build compiles: `tsc --noEmit` ✓ and `vite build` ✓
- [x] Update `src/commands/command-menu.ts` — Anchor + textarea queries updated for pi-input
- [x] Update `src/commands/builtins.ts` — References updated from agent-interface → pi-sidebar, removed dead #empty-state queries

## Summary of changes
- **New**: `pi-sidebar.ts` (~190 lines), `pi-input.ts` (~130 lines)
- **Rewritten**: `theme.css` (12 clean sections, !important down from 30+ to 13), `taskpane.ts` (~220 lines shorter)
- **Updated**: `index.ts`, `command-menu.ts`, `builtins.ts`
- **Untouched**: `dialog.ts`, `boot.ts`, all tools/auth/context/extensions
- **Architecture**: 5 fighting style layers → 2 clean layers (pi-web-ui app.css + our theme.css)
