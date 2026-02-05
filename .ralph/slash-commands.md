# Slash Command System

## Requirements
1. When user types `/` as first char in textarea, show a filterable command menu
2. Built-in commands: /settings, /model, /default-models, /copy, /name, /shortcuts, /share-session
3. Extensible: registry for skills, extensions, prompt snippets
4. Menu appears above textarea, filters as user types, arrow keys + enter to select
5. ESC dismisses menu

## Checklist
- [ ] Create `src/commands/` module with SlashCommand type + registry
- [ ] Built-in commands: settings, model, default-models, copy, name, shortcuts, share-session
- [ ] Create `src/ui/command-menu.ts` — the popup menu component (absolute positioned above textarea)
- [ ] Wire to textarea: detect `/` at start, show menu, filter on input, arrow/enter/esc
- [ ] CSS for command menu (frosted glass, matches design)
- [ ] Extension point: `registerCommand(name, handler)` for plugins
- [ ] Implement each command's action
- [ ] Build + screenshot verification
