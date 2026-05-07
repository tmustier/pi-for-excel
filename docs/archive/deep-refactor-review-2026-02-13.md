# Deep Refactor Review (2026-02-13)

## Executive summary

The codebase is in good functional shape (`npm run check` passes), but complexity is highly concentrated in a small set of files. The strongest refactor opportunities are architectural, not style-level.

Biggest opportunities:

1. **Split the workbook recovery subsystem** (`src/workbook/recovery-states.ts`, `src/workbook/recovery-log.ts`) into focused modules.
2. **Standardize workbook mutation tool pipelines** (`write_cells`, `fill_formula`, `python_transform_range`, `format_cells`, `modify_structure`, `comments`, etc.) to remove repeated audit/recovery scaffolding.
3. **Decompose `taskpane/init.ts`** (composition root + runtime refresh orchestration) into smaller services.
4. **Modularize UI tool cards/humanizers** (`src/ui/tool-renderers.ts`, `src/ui/humanize-params.ts`) into per-tool registries.
5. **Consolidate bridge/server duplication** across `tmux`/`python`/`libreoffice` tools and scripts.

---

## Method used for this review

- Structural scan: file-size and hotspot analysis across `src/`, `tests/`, `scripts/`.
- Coupling scan: import fan-out/fan-in spot checks.
- Duplication scan: repeated helper names and repeated bridge/tooling patterns.
- Validation run:
  - `npm run check` ✅
  - `npm run build` ✅ (but with large chunk warnings + dynamic import/static import overlap warnings)

---

## Hotspots (by size)

| File | LOC |
|---|---:|
| `src/workbook/recovery-states.ts` | 3428 |
| `src/extensions/sandbox-runtime.ts` | 1768 |
| `src/taskpane/init.ts` | 1403 |
| `src/workbook/recovery-log.ts` | 1240 |
| `src/ui/tool-renderers.ts` | 1162 |
| `src/files/workspace.ts` | 1086 |
| `src/ui/humanize-params.ts` | 986 |
| `scripts/python-bridge-server.mjs` | 984 |
| `scripts/tmux-bridge-server.mjs` | 950 |
| `src/tools/tool-details.ts` | 862 |
| `src/tools/mcp.ts` | 833 |
| `src/commands/builtins/extensions-overlay.ts` | 821 |
| `src/ui/files-dialog.ts` | 809 |
| `src/commands/builtins/integrations-overlay.ts` | 750 |
| `src/ui/pi-sidebar.ts` | 736 |
| `src/commands/builtins/experimental.ts` | 735 |

Concentration is strongest in: **workbook recovery**, **runtime/bootstrap**, **tool card UI**, **extensions**, and **bridge integrations**.

---

## Detailed refactor opportunities

## 1) Break up workbook recovery into domain modules (highest leverage)

**Files:**
- `src/workbook/recovery-states.ts` (3428 LOC)
- `src/workbook/recovery-log.ts` (1240 LOC)
- `tests/workbook-recovery-log.test.ts` (1403 LOC)

**Why this matters**
- `recovery-states.ts` mixes several domains in one file: format-state capture/apply, structure-state capture/apply, conditional-format handlers, comment state, cloning/guards/utilities.
- `recovery-log.ts` mixes persistence codec, storage integration, append/list semantics, and restore strategy orchestration.
- Very high cognitive load and high regression risk when touching any recovery behavior.

**Refactor direction**
- Split into modules with a thin facade:
  - `src/workbook/recovery/types.ts`
  - `src/workbook/recovery/format-state.ts`
  - `src/workbook/recovery/structure-state.ts`
  - `src/workbook/recovery/conditional-format-state.ts`
  - `src/workbook/recovery/comment-state.ts`
  - `src/workbook/recovery/codec.ts` (parse/serialize/versioned payload)
  - `src/workbook/recovery/log.ts` (append/list/restore orchestration only)
- Keep public API stable through `src/workbook/recovery-states.ts` + `src/workbook/recovery-log.ts` re-export adapters first.

**Expected impact**: very high maintainability gain, easier testing and safer future recovery features.

---

## 2) Extract a shared mutation-tool pipeline

**Files (main examples):**
- `src/tools/write-cells.ts`
- `src/tools/fill-formula.ts`
- `src/tools/python-transform-range.ts`
- `src/tools/format-cells.ts`
- `src/tools/modify-structure.ts`
- `src/tools/comments.ts`

**Why this matters**
Repeated patterns appear across mutation tools:
- blocked-state handling
- audit append
- recovery checkpoint append + event dispatch
- `appendResultNote(...)` behavior
- similar error/result scaffolding

This makes behavior drift likely (one tool gets a fix, others miss it).

**Refactor direction**
Create shared helpers in e.g. `src/tools/mutation/`:
- `finalizeMutationResult(...)` (audit + recovery checkpoint + note behavior)
- `buildBlockedResult(...)`
- shared `MutationOutcome` typing
- optional per-tool hooks (custom summary text, custom details payload)

**Expected impact**: high consistency + lower bug risk for future tool changes.

---

## 3) Decompose `taskpane/init.ts` into lifecycle services

**File:** `src/taskpane/init.ts` (1403 LOC, 62 imports)

**Why this matters**
`initTaskpane()` currently owns:
- boot/auth/storage setup
- runtime creation and capability refresh logic
- tab/session orchestration
- command wiring
- sidebar event wiring
- status/popover wiring
- focus/visibility refresh policy

This is a composition root plus multiple domain services in one function.

**Refactor direction**
Split into focused units:
- `taskpane/bootstrap-runtime.ts` (agent/runtime creation)
- `taskpane/runtime-capabilities.ts` (tools/system prompt refresh policy)
- `taskpane/sidebar-wiring.ts`
- `taskpane/commands-wiring.ts`
- `taskpane/statusbar-wiring.ts`
- `taskpane/workbook-refresh.ts`

Keep `init.ts` as orchestration glue.

**Expected impact**: high readability and safer changes around session/tab/runtime behaviors.

---

## 4) Modularize tool card rendering and humanizers

**Files:**
- `src/ui/tool-renderers.ts` (1162 LOC)
- `src/ui/humanize-params.ts` (986 LOC)

**Why this matters**
Both files are large dispatch centers with many tool-specific branches. They are now merge-conflict hotspots whenever a tool is added/changed.

**Refactor direction**
- Introduce per-tool modules:
  - `src/ui/tool-renderers/<tool>.ts`
  - `src/ui/humanizers/<tool>.ts`
- Keep a typed registry map in one index file.
- Preserve `TOOL_NAMES_WITH_RENDERER` / `TOOL_NAMES_WITH_HUMANIZER` as source-of-truth gates.

**Expected impact**: medium-high maintainability gain, lower merge conflict rate.

---

## 5) Overlay architecture consolidation

**Files:**
- `src/commands/builtins/extensions-overlay.ts` (821 LOC)
- `src/commands/builtins/integrations-overlay.ts` (750 LOC)
- `src/commands/builtins/overlays.ts` (677 LOC)
- `src/ui/files-dialog.ts` (809 LOC)

**Why this matters**
These overlays share a lot of imperative DOM construction patterns and local helper duplication (`createButton`, `createInput`, `createSectionTitle`, etc.), but evolve independently.

**Refactor direction**
- Keep `overlay-dialog.ts` as lifecycle primitive, add a small shared component helper layer:
  - `src/ui/overlay-kit.ts` (fields/buttons/sections/status rows)
- Split each large overlay into feature sections (installed list, permissions panel, bridge settings, etc.)
- Keep behavior unchanged while reducing single-file size.

**Expected impact**: medium maintainability + easier UI consistency updates.

---

## 6) Unify bridge-tool client logic (`tmux` / `python_run` / `libreoffice_convert`)

**Files:**
- `src/tools/tmux.ts`
- `src/tools/python-run.ts`
- `src/tools/libreoffice-convert.ts`
- `src/tools/python-transform-range.ts` (adjacent integration path)

**Why this matters**
There is clear repeated logic in bridge tools (param parsing patterns, timeout/error extraction, bridge URL/token resolution). This has already started diverging.

**Refactor direction**
- Add `src/tools/bridge-client/` shared primitives:
  - URL/token/settings resolution
  - timeout + fetch wrappers
  - normalized bridge error extraction
  - reusable response parsing helpers
- Keep tool-specific schemas/results separate.

**Expected impact**: medium-high consistency and less repeated bug-fix work.

---

## 7) Consolidate bridge server script infrastructure

**Files:**
- `scripts/python-bridge-server.mjs` (984 LOC)
- `scripts/tmux-bridge-server.mjs` (950 LOC)
- `scripts/cors-proxy-server.mjs` (457 LOC)

**Why this matters**
Repeated server scaffolding exists across scripts (HTTPS toggle, allowed origins, loopback checks, bearer-token auth, JSON body parsing, response helpers). Security-sensitive duplication raises drift risk.

**Refactor direction**
- Create `scripts/server-lib/` shared modules:
  - common server bootstrap (HTTP/HTTPS/cert loading)
  - CORS allowlist handling
  - auth helper
  - request body parsing + error utilities
- Keep endpoint-specific behavior in each server entrypoint.

**Expected impact**: high security maintainability and smaller script footprint.

---

## 8) Remove dynamic-import/static-import overlap for app storage and improve chunking

**Evidence from build**
- Large output: `dist/assets/taskpane-*.js` ≈ **2.5 MB** minified.
- Vite warnings show many modules are both dynamically and statically imported, preventing expected chunk separation (notably `@earendil-works/pi-web-ui/dist/storage/app-storage.js`).

**Files involved (examples):**
- `src/taskpane/init.ts` (static storage path)
- multiple dynamic loaders (`src/workbook/recovery-log.ts`, `src/tools/mcp.ts`, `src/tools/python-run.ts`, `src/tools/web-search.ts`, `src/files/workspace.ts`, etc.)

**Refactor direction**
- Introduce a single storage access facade (`src/storage/get-app-storage.ts`) with explicit policy.
- Prefer dependency injection from init/runtime where possible over repeated dynamic imports.
- Revisit lazy boundaries so imports are either intentionally static or intentionally lazy (not both).

**Expected impact**: medium performance + cleaner loading model.

---

## 9) Extension runtime layering

**Files:**
- `src/extensions/sandbox-runtime.ts` (1768 LOC)
- `src/extensions/runtime-manager.ts` (646 LOC)
- `src/commands/extension-api.ts` (665 LOC)

**Why this matters**
Runtime manager, sandbox RPC transport, and extension API surface are tightly coupled. The sandbox runtime file alone mixes messaging protocol, iframe source generation, UI surface control, and action routing.

**Refactor direction**
Split into layers:
- `extensions/sandbox/protocol.ts` (envelopes + validation)
- `extensions/sandbox/transport.ts` (request/response + timeouts)
- `extensions/sandbox/surfaces.ts` (overlay/widget rendering bridge)
- `extensions/sandbox/host.ts` (activation wiring)

**Expected impact**: medium-high maintainability, easier targeted testing.

---

## 10) Files workspace service split

**Files:**
- `src/files/workspace.ts` (1086 LOC)
- `src/ui/files-dialog.ts` (809 LOC)

**Why this matters**
`workspace.ts` currently owns backend selection, metadata tagging, audit, persistence, and eventing in one service. `files-dialog.ts` is similarly large and stateful.

**Refactor direction**
- Split service internals:
  - `files/workspace-backend-selector.ts`
  - `files/workspace-metadata-store.ts`
  - `files/workspace-audit-store.ts`
  - `files/workspace-service.ts`
- For dialog: split render/update sections (list, viewer, audit, filter state).

**Expected impact**: medium maintainability + easier feature additions.

---

## Quick wins (low risk, immediate)

1. Replace local `getErrorMessage` duplicates with `src/utils/errors.ts` (currently defined in multiple files).
2. Deduplicate `formatRelativeDate` (`builtins/overlays.ts`, `ui/files-dialog.ts`) into shared util.
3. Deduplicate bridge URL/settings access helpers shared by:
   - `commands/builtins/experimental.ts`
   - `tools/experimental-tool-gates/evaluation.ts`
4. Consolidate shared mutation note helper (`appendResultNote`) into one place.
5. Split `tests/workbook-recovery-log.test.ts` by snapshot kind to reduce single-file blast radius.

---

## Suggested phased roadmap

### Phase 1 (highest ROI, low behavior risk)
- Recovery subsystem split (state + log + codec modules)
- Shared mutation pipeline extraction

### Phase 2 (developer velocity)
- `taskpane/init.ts` decomposition
- Tool-renderer/humanizer modularization

### Phase 3 (platform hardening)
- Bridge tool + server shared infrastructure
- Extension sandbox/runtime layering split

### Phase 4 (performance and bundle hygiene)
- Storage import strategy cleanup
- Lazy-boundary cleanup + chunking pass

---

## Final note

Earlier refactor docs in this repo (e.g. `docs/refactor-plan.md`) indicate a prior modularization pass that succeeded. This review finds that feature growth has since re-concentrated complexity. A second structured refactor pass is warranted now, especially around recovery, runtime orchestration, and integration tooling.
