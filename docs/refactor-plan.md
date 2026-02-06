# Refactor Plan — modularize `taskpane` + `builtins`

Context: codebase grew from rapid prototyping. Lint + typecheck are now clean; next priority is **maintainability** (smaller modules, clearer ownership) with **no behavior changes**.

## Goals

- Reduce “god files” (target: **≤ ~400–500 LOC** per file where practical).
- Make responsibilities obvious:
  - taskpane bootstrap / agent wiring
  - session persistence
  - context injection (blueprint + selection + change tracking)
  - slash command registration
- Keep type-safety improvements intact (no re-introducing `any`/`!`).
- Keep runtime behavior the same (refactor-only).

## Non-goals (for this pass)

- No tool behavior changes (see `src/tools/DECISIONS.md`).
- No UI redesign.
- No dependency upgrades.
- No new storage formats / migrations.

## Baseline / guardrails

Before and after each step:

- `npm run check` (lint + typecheck)
- `npm run test:models`
- `npm run build`
- `npm run validate`

Prefer small, reversible commits.

---

## Phase 1A — extract shared utilities (reduce duplication)

### 1) `src/utils/type-guards.ts`

Unify helpers currently duplicated across files:

- `isRecord(value: unknown): value is Record<string, unknown>`

(Keep it tiny; add more only if duplicated ≥2 places.)

### 2) `src/utils/content.ts`

Unify “message content → text” logic used by `taskpane.ts` and `builtins.ts`:

- `extractTextBlocks(content: unknown): string`
- `extractTextFromContent(content: unknown): string` (string | blocks)
- `summarizeContentForTranscript(content: unknown, limits?)` (move from builtins)

Then replace local copies in:
- `src/taskpane.ts`
- `src/commands/builtins.ts`
- (optionally) other call sites that do the same thing

**Definition of done:** no change in exported transcript format / session preview text.

---

## Phase 1B — split `src/taskpane.ts` (1115 LOC)

Target shape:

```
src/taskpane/
  dom.ts
  init.ts
  model-selector-patch.ts
  sessions.ts
  queue-display.ts
src/taskpane.ts   (thin entrypoint)
```

### 1) `src/taskpane/model-selector-patch.ts`

Move the ModelSelector private-method patch out of `taskpane.ts`.

- Keep behavior identical:
  - filter providers based on configured API keys
  - keep current model at top
  - featured model ordering rules unchanged
- Add a runtime guard:
  - if `getFilteredModels` is missing, log a warning and no-op (fail soft)

Expose a single initializer:
- `installModelSelectorPatch()`
- and a setter for active providers (if still needed): `setActiveProviders(providers: Set<string>)`

### 2) `src/taskpane/dom.ts`

Move DOM helpers:

- `getRequiredElement()`
- error banner helpers: `showErrorBanner`, `clearErrorBanner`

### 3) `src/taskpane/sessions.ts`

Extract session persistence from inside `init()` into a focused module.

Proposed API:

- `setupSessionPersistence({ agent, sessions, sidebar, extractTextFromContent })`
  - returns `{ startNewSession, setSessionTitle, adoptResumedSession }`

Also extract pure helpers:
- `computeSessionUsage(messages)`
- `buildSessionPreview(messages)`

**Key constraint:** preserve the metadata shape expected by `SessionsStore.saveSession()`.

### 4) `src/taskpane/queue-display.ts`

Extract queued steering/follow-up UI wiring:

- maintain current DOM implementation (no Lit rewrite yet)
- provide:
  - `createQueueDisplay(sidebar)` → `{ add, clear, syncFromAgentEvent? }`

### 5) `src/taskpane/init.ts`

Move the bulk of `init()` (storage, auth restore, agent creation, sidebar mount, subscriptions, event listeners).

`src/taskpane.ts` becomes the thin entrypoint that:

- `import "./boot.js"` first
- installs patches
- runs `Office.onReady(...)` + 3s fallback
- calls `initTaskpane()` from `src/taskpane/init.ts`

**Definition of done:** identical startup behavior (incl. 3s Office fallback), same events (`pi:*`), same status bar updates.

---

## Phase 1C — split `src/commands/builtins.ts` (612 LOC)

Target shape:

```
src/commands/builtins/
  index.ts
  model.ts
  settings.ts
  session.ts
  export.ts
  clipboard.ts
  help.ts
src/commands/builtins.ts  (optional shim or removed)
```

Approach:

- Each file exports a `registerXCommands(agent)` function.
- `src/commands/builtins/index.ts` exports the public `registerBuiltins(agent)`.
- Move helper functions (transcript serialization, block formatting) to `src/utils/content.ts` where shared.

**Definition of done:**
- commands list unchanged (names/descriptions)
- command side-effects unchanged (dialogs, events, toasts)
- `npm run check` stays clean

---

## Phase 1D — follow-ups (optional, after Phase 1)

Not required for the first pass, but likely next:

- Quarantine monkey patches into a clear `src/compat/*` namespace (`boot.ts` and ModelSelector patch)
- Reduce remaining `as unknown as` to one documented helper (tool registry invariance)
- Blueprint cache invalidation strategy: decide which tools call `invalidateBlueprint()`

---

## Milestone checklist

- [ ] Phase 1A merged (shared utils)
- [ ] Phase 1B merged (taskpane modularized)
- [ ] Phase 1C merged (builtins modularized)
- [ ] Verify in Excel manually (sideload + basic chat, model switch, login, session restore)
