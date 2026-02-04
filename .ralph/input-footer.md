# Input Area + Footer Redesign

## Issues from User
1. Model picker doesn't persist selection 
2. Thinking level picker doesn't change
3. At narrow widths: send arrow disappears, model name overflows
4. Footer should show: context % (e.g. 49%/200k) and thinking level with shift+tab shortcut
5. Textarea max-height too small (~2.5 lines) — should grow much taller like Claude

## Checklist
- [ ] Fix textarea max-height (increase to ~50vh for proper multi-line)
- [ ] Debug model picker persistence (check if setModel() is called)
- [ ] Debug thinking level selector
- [ ] Hide attachment button (not needed for Excel sidebar)
- [ ] Compact model display — truncate or move out of cramped toolbar
- [ ] Add shift+tab keyboard shortcut to cycle thinking levels
- [ ] Add custom status bar showing context % and thinking level
- [ ] Test at narrow widths (320px, 375px)
- [ ] Build + screenshot verification
