# AGENTS.md

**Last reviewed:** 2026-02-12

Notes for agents working in this repo.

## Read before changing behavior
- Tool behavior rules: `src/tools/DECISIONS.md`
- UI/CSS architecture: `src/ui/README.md` (Tailwind v4 `@layer` gotcha)
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

### Tool results (`ToolResultMessage.details`)
- Keep human-readable output in `result.content`.
- Put stable machine metadata in `result.details`.
- UI should prefer `details`, with fallback for older persisted sessions.
- Reuse guards/types from `src/tools/tool-details.ts`.

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
- Avoid explicit `any` / `as any`; prefer specific types, unions, generics, or `unknown` + narrowing.
- Avoid non-null assertions where practical; use guards/early throws.

## Verification
- `npm run check`
- `npm run build`
- `npm run test:models`
- `npm run test:context` when touching prompt/context/tool disclosure/session wiring
- `npm run test:security` when touching proxy/bridge/auth/HTML safety paths
- Manual Excel smoke test when touching session persistence, tools, auth, or UI wiring
- For UI/bootstrap issues (e.g. endless "Initializing…"), inspect the page with `agent-browser` first:
  - `agent-browser --ignore-https-errors open https://localhost:3000/src/taskpane.html`
  - `agent-browser snapshot -i -c --json`
  - `agent-browser console --json`
  - `agent-browser errors --json`

## Pre-commit
- `.githooks/pre-commit` runs `npm run lint` + `npm run typecheck`.
- Bypass only when needed: `git commit --no-verify`.

## Excel sideloaded manifest gotcha (macOS)
Excel loads a sideloaded manifest from:
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/{add-in-id}.manifest.xml`

If local changes do not show up:
1. Verify sideloaded manifest points to `https://localhost:3000/...` (not production URL).
2. Recopy manifest:
   `cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/a1b2c3d4-e5f6-7890-abcd-ef1234567890.manifest.xml`
3. Quit Excel fully and reopen.
4. If still stale, clear WKWebView cache and relaunch Excel:
   - `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/`
   - `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/WebKit/`
