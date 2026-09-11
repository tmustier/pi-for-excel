# AGENTS.md

**Last reviewed:** 2026-02-12

Notes for agents working in this repo.

## Read before changing behavior
- Agent coding standards: `docs/coding-standards.md`
- Tool behavior rules: `src/tools/DECISIONS.md`
- UI/CSS architecture: `src/ui/README.md` (first-party preflight; keep it faithful)
- Upstream divergences: `docs/upstream-divergences.md` (read before adding new divergences)
- Docs index: `docs/README.md`
- Model registry freshness: `docs/model-updates.md` (if **Last verified** > 1 week, refresh Pi deps + re-verify model IDs before model UX changes)

## High-leverage conventions

### Core tools: one source of truth
- Define core tool names in `src/tools/registry.ts` (`CORE_TOOL_NAMES`, `CoreToolName`, `createCoreTools()`).
- Do not duplicate tool-name lists; import `CORE_TOOL_NAMES`.
- When adding/removing a core tool, update in the same PR:
  - `src/tools/registry.ts`
  - `src/ui/tool-renderers.ts`
  - `src/ui/humanize-params.ts`
  - `src/context/tool-disclosure.ts`
  - `src/prompt/system-prompt.ts` (if documented tool list changes)

### UI i18n (`t()` layer)
- All user-visible UI strings go through `t()` from `src/language/index.ts`.
  `en.json` is the source of truth (values must exactly match the intended
  English UI); `zh-CN.json` mirrors the key set. Parity + usage enforced by
  `tests/i18n-locales.test.ts` (runs in `test:context`).
- **Never call `t()` at module scope** — language is set at boot, after module
  import; module-scope calls freeze to English. Store keys and resolve lazily
  (see `src/ui/whimsical-messages.ts`).
- **Never route agent-facing strings through `t()`**: system prompt, tool
  names/descriptions/schemas, context injection, compaction prompts. Tool
  `label` is UI-only and safe. See prompt-caching notes below.
- zh-CN is AI-generated; keep the "AI 翻译" marker on the language option
  (`settings.section.language.zh`).

### Tool results (`ToolResultMessage.details`)
- Keep human-readable output in `result.content`.
- Put stable machine metadata in `result.details`.
- UI should prefer `details`, with fallback for older persisted sessions.
- Every details payload is a TypeBox schema in `src/tools/tool-details.ts` with its type as `Static<>`; add a new payload there and to `toolDetailsSchemasByKind`. The UI decodes `result.details` once with `decodeToolDetails()` and switches on `kind`; do not add per-kind `isXDetails` guards.

### Workbook identity + session restore
- Never persist raw `Office.context.document.url`.
- Use `getWorkbookContext()` from `src/workbook/context.ts`.
- Use `src/workbook/session-association.ts` helpers for SettingsStore mapping keys.

### Security / HTML / local servers
- Avoid `innerHTML` for user/tool/session content; use DOM APIs or `src/utils/html.ts`.
- Keep markdown protections from `installMarkedSafetyPatch()` (`src/compat/marked-safety.ts`).
- Keep strict origin allowlists in:
  - `scripts/cors-proxy-server.mjs`
  - `scripts/tmux-bridge-server.mjs`
  - `scripts/python-bridge-server.mjs`
- Keep proxy target filtering strict in `scripts/proxy-target-policy.mjs`.
  - Do not commit permissive defaults (e.g. `ALLOW_ALL_TARGET_HOSTS=1`).

### Bundle hygiene (Office WebView)
- Avoid Node-only imports and side-effect barrel imports.
- After import/dependency changes, run `npm run build` and check chunk sizes + Vite browser-compat warnings.

### Prompt caching gotchas
- Prompt cache keys are prefix-based and sensitive to: model identity, system prompt, tool schemas, and session key.
- Keep static prefix content stable (no timestamps/random IDs in system prompt or tool metadata).
- Prefer message-tail updates for volatile state (auto-context/system reminders) instead of mutating base prompt text every turn.
- Keep tool ordering deterministic; do not rebuild tool lists with unstable ordering.
- Do not reintroduce blanket eager `setTools(...)` on refresh passes when extension tools exist; use fingerprint + extension tool revision semantics.
- When changing context/tool/model wiring, validate against `docs/cache-observability-baselines.md` and record expected vs observed `prefixChangeReasons`.

## TypeScript policy
- No `// @ts-ignore`.
- If unavoidable: `// @ts-expect-error -- <reason>` with a real reason.
- No explicit `any`, `as any`, or non-null assertions.
- `unknown` only at the seam (JSON decode, host object, bridge/sandbox message, caught error); parse it into a concrete domain type on the next line. Never add `unknown` further in: `npm run check:unknown-burndown` fails if the count rises (see `docs/coding-standards.md`).
- Do not add generic object/record guards (`isRecord`, `isObjectValue`, `isPlainObject`, `is…PayloadShape`, etc.); an object check alone does not establish a domain contract. Parse the domain shape, preferably with a TypeBox schema and `Static<>`.
- Schemas use `typebox` (pinned to pi-ai's version; one copy ships). Import `StringEnum` from `src/tools/string-enum.ts` (pi-ai's, with a `const` type parameter so inline arrays keep literal types).
- Type assertions and lint suppressions must stay local and explain the safety invariant; do not add blanket safety comments.

## Verification

Use the local acceptance policy in `docs/coding-standards.md` and the `.agents/skills/excel-background-verification/SKILL.md` workflow.

The Excel lane runs entirely in the background, including opening the pane after a relaunch (the skill bootstraps it from the scratch workbook) and reading what the taskpane persisted (straight from WebKit's IndexedDB). Work through the skill's "Before you call it blocked" table before reporting the lane blocked; each row there has been hit and worked around.

- `npm run check`
- `npm run build`
- `npm run test:models`
- `npm run test:context` when touching prompt/context/tool disclosure/session wiring
- `npm run test:security` when touching proxy/bridge/auth/HTML safety paths
- Manual Excel smoke test when touching session persistence, tools, auth, or UI wiring

### Visual UI verification (agent-browser)

Use the **UI Gallery** (`src/ui-gallery.html`) to verify CSS and component changes
without needing Excel. It renders mock components with the real CSS theme.

```bash
# Ensure dev server is running (starts automatically if needed)
./scripts/ui-verify.sh                    # Full gallery screenshot
./scripts/ui-verify.sh diff-table         # Screenshot a specific section
./scripts/ui-verify.sh taskpane           # Screenshot the real taskpane (waits for Office timeout)
./scripts/ui-verify.sh stop               # Clean up browser session
```

Available gallery sections (use as argument):
`badges`, `file-items`, `tool-cards`, `tool-groups`, `diff-table`,
`text-preview`, `buttons`, `toasts`, `markdown`

Or use agent-browser directly for interactive inspection:
```bash
npx agent-browser --session pi-ui open http://localhost:3141/src/ui-gallery.html
npx agent-browser --session pi-ui wait 2000
npx agent-browser --session pi-ui snapshot -i          # See interactive elements
npx agent-browser --session pi-ui screenshot shot.png  # Capture
npx agent-browser --session pi-ui close                # Clean up
```

For **full taskpane** inspection (boots without Excel after 3s timeout):
```bash
npx agent-browser open http://localhost:3141/src/taskpane.html
npx agent-browser wait 4000           # Wait for Office.js fallback
npx agent-browser snapshot -i -c      # Interactive snapshot
npx agent-browser console --json      # Check for JS errors
npx agent-browser errors --json       # Check for page errors
```

When to add new gallery sections:
- Adding a new component type → add a mock render in `src/ui-gallery.ts`
- Changing CSS for an existing component → verify via `./scripts/ui-verify.sh <section>`
- Before/after comparison → screenshot before change, make edit, screenshot again

## Pre-commit
- `.githooks/pre-commit` runs `npm run check`.
- Bypass only when needed: `git commit --no-verify`.

## Excel sideloaded manifest gotcha (macOS)
Excel loads a sideloaded manifest from:
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/{add-in-id}.manifest.xml`

If local changes do not show up:
1. Verify sideloaded manifest points to `https://localhost:3141/...` (not production URL).
2. Recopy manifest:
   `cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/a1b2c3d4-e5f6-7890-abcd-ef1234567890.manifest.xml`
3. Quit Excel fully and reopen.
4. If still stale, clear WKWebView cache and relaunch Excel:
   - `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/`
   - `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/WebKit/`

No `Open Pi` ribbon button after a relaunch is expected, not a stale manifest: Excel keeps sideloaded add-ins out of the ribbon cache until they have been activated once in the running process. Open the pane from a workbook that carries the auto-open part (`.agents/skills/excel-background-verification/scripts/scratch-workbook.py`); the button appears after that. Clearing the WebKit cache also deletes the taskpane's IndexedDB (sessions, recovery snapshots, files-workspace state), so do it last.
