# Input Area + Footer Redesign

## Issues from User
1. Model picker doesn't persist selection 
2. Thinking level picker doesn't change
3. At narrow widths: send arrow disappears, model name overflows
4. Footer should show: context % (e.g. 49%/200k) and thinking level with shift+tab shortcut
5. Textarea max-height too small (~2.5 lines) — should grow much taller like Claude

## Checklist
- [x] Fix textarea max-height (increased to 50vh)
- [x] Hide attachment button (not needed for Excel sidebar)
- [x] Hide thinking selector from toolbar (replaced by status bar)
- [x] Compact model display — truncated responsively with min(140px, calc(100vw-180px))
- [x] Send button: flex-shrink:0, min-width:26px — never disappears
- [x] Add shift+tab keyboard shortcut to cycle thinking levels
- [x] Add custom status bar: context % / window size (left) + brain icon + thinking level (right)
- [x] Status bar thinking indicator is clickable (cycles levels)
- [x] Lucide Brain SVG icon (matches pi-web-ui)
- [x] Test at narrow widths (320px, 375px) — verified
- [x] Build + screenshot verification
- [x] Model picker persistence — verified: setModel() works, issue was likely narrow-width UI
- [x] Thinking level — verified: setThinkingLevel() syncs with agent state + requestUpdate()
