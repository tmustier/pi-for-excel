# UI Ownership: dropping pi-web-ui and owning the rendering layer

**Status:** In progress (branch `ui-ownership`)
**Date:** 2026-07-07
**Decision:** Own the entire UI layer first (Phase 1), then overhaul UX/visuals on the owned foundation (Phase 2).

---

## 1. Why

### pi-web-ui is dead upstream

- `@earendil-works/pi-web-ui` last published **0.75.3 on 2026-05-18** (plus a
  0.74.2 backport on 2026-05-21). Nothing since.
- The `web-ui` package has been **removed from the pi-mono monorepo entirely**
  (`packages/` now contains only `agent`, `ai`, `coding-agent`, `orchestrator`,
  `tui`). There will never be another release.
- Meanwhile `@earendil-works/pi-agent-core` / `@earendil-works/pi-ai` are at
  **0.80.3 (2026-06-30)** and actively maintained. The agent runtime dependency
  is healthy — only the UI layer died.
- We are already carrying compat shims to bridge the skew:
  - `vite.config` alias mapping `@earendil-works/pi-ai` → `/compat` for
    pi-web-ui's dist modules
  - `src/compat/lit-class-field-shadowing.ts`
  - `src/compat/model-selector-patch.ts`
  - `src/compat/marked-safety.ts` (partially ours regardless)

The original bet ("anchor on pi-web-ui to get upstream UI evolution for free")
is falsified. Continuing to build on it means designing every future pixel
twice.

### We already own ~90% of the UI

Our own code: `pi-sidebar.ts` (1.2k lines), `pi-input.ts`, `tool-renderers.ts`
(1.3k lines), tabs, status bar, ~26 slash commands, every overlay/dialog
(settings, extensions hub, files, recovery, resume, rules, shortcuts…),
toasts, ~7k lines of theme CSS.

What we actually consume from pi-web-ui (audited 2026-07-07):

| Area | Modules | Compiled size |
|---|---|---|
| Rendering | `Messages.js` (User/Assistant/Tool/Aborted + `defaultConvertToLlm`), `MessageList`, `StreamingMessageContainer`, `AttachmentTile`, `message-renderer-registry`, `tools/renderer-registry`, `tools/types` | ~1,100 lines |
| Dialogs | `ModelSelector`, `ApiKeyPromptDialog` | ~450 lines |
| Storage | `app-storage`, `store`, `types`, `backends/indexeddb-storage-backend`, `stores/{settings,sessions,provider-keys,custom-providers}` | ~430 lines |
| Utils | `utils/proxy-utils` | tiny |
| CSS | `app.css` (Tailwind v4 foundation) | 85 KB |

From `@mariozechner/mini-lit` (also stale since 2026-04-16): `icon`/`iconDOM`
helpers, `MarkdownBlock` (256 lines), `CodeBlock` (89 lines).

### The crust is actively expensive

An entire apparatus exists purely to style someone else's Tailwind-in-Light-DOM
internals:

- `message-style-hooks.ts` / `dialog-style-hooks.ts` (stamp semantic classes
  onto pi-web-ui DOM after render)
- `theme/unstable-overrides.css` (utility-coupled upstream selectors)
- The Tailwind `@layer` gotcha (unlayered CSS silently beats utilities)
- `check:theme-utility-overrides` lint check
- Six `src/stubs/pi-web-ui-*` stubs + a dedicated vite plugin to stop the
  package barrel from dragging artifacts/attachments/REPL code into the bundle
- 195 KB of shipped CSS to render a chat pane

Only **3 of our own files** use Tailwind utility classes in templates
(`tool-renderers.ts`, `pi-sidebar.ts`, `message-renderers.ts`) — exactly the
pi-web-ui-facing ones. Everything else is already semantic-class discipline.
Dropping Tailwind is therefore cheap.

## 2. Options considered

- **A. Fork pi-web-ui wholesale** — unblocks version pinning but inherits the
  Tailwind/Light-DOM architecture we are already fighting. Half-measure.
  Rejected.
- **B. Clean-room the crust; own everything** — write our own message
  components in Lit, vendor the tiny storage layer, drop Tailwind, style from
  our existing token system. **Chosen.**
- **C. Replatform (React/assistant-ui, etc.)** — rewrite of 40+ Lit files for
  no user-visible gain, worse bundle in the Office WebView. Rejected.

## 3. Phase 1 — Ownership migration (this branch)

Goal: zero runtime imports from `@earendil-works/pi-web-ui` and
`@mariozechner/mini-lit`; visual rough-parity; all checks/tests green; smaller
bundle. `pi-agent-core` / `pi-ai` remain as the maintained runtime deps.

Steps (each a coherent commit):

1. **Findings/plan doc** (this file).
2. **Vendor storage** → `src/storage/local/` (port of app-storage, store,
   types, IndexedDB backend, and the four stores; MIT-attributed). Rewire all
   imports. API-identical to minimize churn.
3. **Own icon helper** → replace mini-lit `icon`/`iconDOM` with a small local
   lucide wrapper.
4. **Own markdown/code rendering** → `<pi-markdown>` / `<pi-code-block>`
   using our existing `marked` + `installMarkedSafetyPatch()` pipeline,
   styled with first-party semantic CSS (no Tailwind classes).
5. **Clean-room message layer** → `src/ui/messages/`: `MessageList`,
   `StreamingMessageContainer`, `UserMessage`, `AssistantMessage`,
   `ToolMessage`, `AbortedMessage`, `AttachmentTile`, renderer registries,
   and a local `convertToLlm` default (attachment/artifact filtering logic
   currently imported from `Messages.js`). Fold the style-hook semantics
   directly into templates; delete `message-style-hooks.ts` and rewrite the
   3 utility-class-using files to semantic classes.
6. **Own dialogs** → local model selector + API key prompt; delete
   `compat/model-selector-patch.ts`; fold `dialog-style-hooks.ts` semantics
   into the components.
7. **Drop Tailwind + package** → remove `app.css` from boot, add owned base
   styles, delete stubs + vite plugin + aliases + `pi-ai → compat` alias,
   remove `@earendil-works/pi-web-ui` and `@mariozechner/mini-lit` from
   package.json, retire `unstable-overrides.css`, update lint checks and docs
   (`src/ui/README.md`, `docs/upstream-divergences.md`, AGENTS.md pointers).
8. **Verification** → `npm run check`, full `npm test`, `npm run build`
   (bundle-size comparison), ui-gallery + taskpane screenshots, manual Excel
   smoke via excel-background-verification.

Known risk areas and how they are handled:

- **Streaming re-render performance** — StreamingMessageContainer re-renders
  the in-flight assistant message on deltas; keep the same
  "only the last message re-renders" contract and verify with a long
  streaming session.
- **Scroll anchoring** — preserve MessageList's pin-to-bottom behavior
  (stick when at bottom, don't yank when scrolled up).
- **Markdown incremental rendering** — `<pi-markdown>` re-parses on each
  delta like markdown-block did; acceptable at sidebar scale, verified in
  streaming smoke.
- **Persisted sessions** — vendored stores must read existing IndexedDB
  data unchanged (same DB name, store names, key paths). Verified by
  restoring a pre-migration session.
- **i18n** — all new user-visible strings go through `t()` with keys added
  to `en.json` + `zh-CN.json` (parity enforced by `tests/i18n-locales.test.ts`).
- **Prompt caching** — this migration is UI-only; no changes to system
  prompt, tool schemas, or message conversion semantics beyond moving
  `defaultConvertToLlm` in-repo verbatim.

## 4. Phase 2 — UX/visual overhaul (follow-up on the owned foundation)

Directions agreed for design review (mockups via ui-gallery before build-out):

1. **Activity condensation** — collapse runs of tool cards into a single live
   activity block ("Working… n steps", expandable), instead of one card per
   tool call. Biggest transcript-density win at 350 px.
2. **Status bar consolidation** — one summary chip opening a single popover
   (model · thinking · context · execution mode) instead of four micro-chips.
3. **Calm empty state** — at most one banner; suggestion cards tightened;
   disclosure moved to first-run settings.
4. **Selection-aware context pill** — "Selection: D10:L14" chip above input,
   tap to include in the prompt.
5. **Inline revert affordance** — surface the recovery/audit story on change
   summaries in the transcript, not only behind `/revert`.
6. **Visual identity pass** — own type scale, refined Excel-adjacent palette,
   consistent radii/spacing from `theme/tokens.css`, density tuned for 350 px.
7. **Discoverability** — queued-message chip (Alt+Enter is invisible today),
   subtle hints for thinking-level cycling.

## 5. Upstream relationship after this change

- `@earendil-works/pi-agent-core` and `@earendil-works/pi-ai` remain upstream
  dependencies; `docs/upstream-divergences.md` philosophy still applies to
  agent behavior.
- The UI layer is no longer an upstream divergence concern: pi-web-ui has no
  upstream to diverge from. `check:pi-lockstep` scope narrows to
  agent-core/pi-ai.
- Vendored storage code retains MIT attribution headers pointing at pi-mono.
