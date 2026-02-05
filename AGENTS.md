# AGENTS.md

Notes for agents working in this repo:

- **Tool behavior decisions live in `src/tools/DECISIONS.md`.** Read it before changing tool behavior (column widths, borders, overwrite protection, etc.).
- **UI architecture lives in `src/ui/README.md`.** Read it before touching CSS or components — especially the Tailwind v4 `@layer` gotcha (unlayered resets clobber all utilities).
- **Docs index:** `docs/README.md` (mirrors Pi's docs layout).
- **Model registry freshness:** check `docs/model-updates.md` → if **Last verified** is > 1 week ago, update Pi deps + re-verify pinned model IDs before changing model selection UX.
