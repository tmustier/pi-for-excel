# Codebase Simplification Plan (Pi for Excel)

Goal: make the repo **simple, easy to understand, easy to navigate, and easy to build on** — while keeping behavior stable.

This plan focuses on a small set of “big levers” that remove ongoing maintenance tax (drift, duplication, brittle parsing) rather than lots of local micro-refactors.

> Context strategy and caching guardrails now live in: [`docs/context-management-policy.md`](../context-management-policy.md).
> Keep simplification work aligned with that policy (especially deterministic tool metadata, context shaping, and debug observability).
>
> Phase 1 recovery/mutation execution tracker (historical; completed on 2026-02-13): [`docs/archive/refactor-execution-plan-2026-02-13.md`](./refactor-execution-plan-2026-02-13.md).
> CLOSE1 is complete: `tests/workbook-recovery-log.test.ts` has been split into:
> - `tests/recovery-log-persistence.test.ts`
> - `tests/recovery-log-restore.test.ts`
> - `tests/recovery-log-format.test.ts`
> - `tests/recovery-log-structure.test.ts`
> - run them with `npm run test:recovery` (Node v25-safe via the repo loader).

---

## 0) Current snapshot (what we observed)

### Strengths
- ✅ `npm run check` (eslint + tsc) is clean.
- ✅ The project is already modularized (no giant `taskpane.ts` god-file anymore).
- ✅ Most modules have clear ownership (`src/tools`, `src/ui`, `src/taskpane`, `src/context`, `src/auth`).

### Hotspots / bloat concentrations (not everywhere)
- `src/ui/humanize-params.ts` (~650 LOC)
- `src/ui/tool-renderers.ts` (~530 LOC)
- `src/tools/format-cells.ts` (~500 LOC)
- `src/ui/provider-login.ts` (~420 LOC)
- CSS: `src/ui/theme/components.css` (~590 LOC), `src/ui/theme/content-overrides.css` (~485 LOC)

### Drift already present (high-signal maintainability smell)
There are multiple tool lists that are already out of sync:
- `src/tools/index.ts` creates **11 tools** (but comment says “10”).
- `src/prompt/system-prompt.ts` documents **11 tools** (includes `comments`).
- `src/ui/tool-renderers.ts`’s `EXCEL_TOOL_NAMES` list omits `comments`.
- `src/ui/humanize-params.ts` registry omits `comments`.
- `src/ui/tool-renderers.ts` still mentions a removed tool (`get_recent_changes`) in `describeToolCall()`.

### Bundling signal (big lever)
`vite build` emits:
- A very large JS chunk and chunk-size warnings.
- Node/browser boundary warnings (e.g. `http` externalized for browser compatibility).

This likely means we’re bundling more of `@earendil-works/pi-web-ui` / `@earendil-works/pi-ai` than the Excel taskpane truly needs.

---

## Roadmap constraints (from open issues)

These constraints should shape cleanup work so we don’t refactor into a dead end:

- **Workbook scoping is coming** (#31, #23, #30). Introduce a `workbookContext` abstraction early, and ensure persistence can be keyed by workbook identity without rewriting stores later.
- **Artifacts/workspace are coming** (#32). File/attachment support pulls in heavy dependencies (PDF/Office parsing) and may require browser filesystem APIs (OPFS / File System Access). Cleanup should favor **lazy-loading** and avoid side-effect imports that defeat tree-shaking.
- **Extensibility is a core product feature** (#13, #24). The registry should be designed for hosted builds:
  - **V1:** paste-code extensions → Blob URL + dynamic `import()` (no `eval`).
  - **V2:** install from URL (GitHub raw / releases) and potentially package sources (npm).

Practical implication: treat the registry work as a *platform registry* refactor, not just a list-of-tools cleanup.

---

## 1) Lever: single source of truth for “Excel tools”

### Problem
Tool names + UI hooks + prompt docs live in multiple places and drift.

### Proposal
Create a single **capability registry** module that owns:
- the canonical list of core tool names (plus a runtime extension registry)
- tool creation (`createXTool()`)
- UI hooks (input humanizers, renderer registration)
- metadata needed for future platform features (tier/progressive disclosure, permissions)

**Sketch:**
- `src/tools/registry.ts`
  - `export const CORE_TOOL_NAMES = [...] as const`
  - `export type CoreToolName = (typeof CORE_TOOL_NAMES)[number]`
  - `export type ToolName = CoreToolName | (string & {})` // extension tools
  - `export function createCoreTools(): AgentTool[]`
  - `export function registerTool(...)` // extensions/integrations
  - `export function getToolUi(...)` // humanizer/renderer hooks

Then:
- `src/tools/index.ts` becomes a thin adapter around the registry.
- UI (`tool-renderers`, `humanize-params`) imports the canonical list/type.

### Immediate quick wins
- Fix the existing drift (add `comments` to renderer/humanizer registration).
- Remove references to removed tools (`get_recent_changes`).
- Fix comments that claim wrong tool count.

### Definition of done
- There is **one** canonical list of Excel tools in the repo.
- Adding/removing a tool requires updating **one** place.
- `npm run check` + `npm run build` still pass.

---

## 2) Lever: structured tool results via `ToolResultMessage.details`

### Problem
The UI currently infers meaning by parsing tool output text (regex-based address extraction, echo detection, error counting). This is brittle and increases the size/complexity of `src/ui/tool-renderers.ts`.

### Proposal
Each tool returns a small, stable metadata payload in `details` (and keeps the human-readable markdown in `content`).

**Examples (minimal):**
- `write_cells.details`:
  - `{ kind: "write_cells", blocked: boolean, address?: string, existingCount?: number, formulaErrorCount?: number }`
- `fill_formula.details`:
  - `{ kind: "fill_formula", blocked: boolean, address?: string, formulaErrorCount?: number }`
- `format_cells.details`:
  - `{ kind: "format_cells", address?: string, warningsCount?: number }`

UI then:
- builds the tool card header from `details` (no regex)
- renders badges (`blocked`, `errorCount`) from structured fields
- keeps markdown rendering for the “Result” body text

### Benefits
- Simplifies `tool-renderers.ts` dramatically.
- Tool output text can evolve without silently breaking UI logic.
- Enables future UI improvements (click-to-navigate addresses, inline previews) without more string parsing.

### Definition of done
- For Excel tools that affect addresses/ranges, the renderer does **not** need regex parsing to:
  - find the written address
  - detect blocked state
  - detect formula error counts

---

## 3) Lever: reduce bundle size + eliminate Node-only leakage

### Problem
The Excel taskpane is a browser webview. Some transitive dependencies pull in Node-only modules or large optional chunks.

A likely major culprit:
- `src/ui/pi-sidebar.ts` does `import "@earendil-works/pi-web-ui";` as a side-effect import to register custom elements.
  - Side-effect imports often defeat tree-shaking and pull in far more than needed.

### Proposal
**Phase A: measurement + containment**
- Keep a baseline record of build output sizes (from `vite build` output).
- Identify the biggest imports (e.g. anything pulling `pdfjs`, `xlsx`, etc.)

**Phase B: targeted reductions**
- Replace side-effect imports with explicit, minimal imports where possible.
- Consider upstream-friendly improvements if needed:
  - a `@earendil-works/pi-web-ui/register-minimal` entrypoint that registers only the message/tool components we use
- Continue stubbing Node-only providers (similar to existing Bedrock + `stream` stubs) for Office builds.

### Definition of done
- `vite build` output JS size decreases meaningfully (or is split into cacheable chunks).
- Fewer “externalized for browser compatibility” warnings.
- No runtime regressions in the taskpane.

---

## 4) Lever: extract an `excel/ops` layer to simplify tool implementations

### Problem
Several tools repeat Office.js ceremony and tricky patterns:
- range address normalization / sheet prefixing
- “load + sync + read back” scaffolding
- cell-in-range logic (duplicated between tools)
- multi-range handling + RangeAreas edge cases

### Proposal
Create a small internal layer:
- `src/excel/ops/*` (or a couple of focused files) that provides:
  - `stripSheet(address)`, `isCellInRange(cell, range)`
  - multi-range iteration helpers
  - common “load, sync, return { sheetName, address }” patterns
  - formatting primitives (borders, numberFormat matrices, etc.)

Then tools become:
1) validate params
2) call `excelOps.*`
3) format a concise markdown result

### Definition of done
- The largest tools (`format_cells`, `read_range`, `comments`) shrink meaningfully.
- Shared range helpers exist in one place.

---

## 5) Lever: standardize overlay UI (provider login, resume, shortcuts)

### Problem
Overlays are built via `innerHTML` with inline styles across:
- `src/ui/provider-login.ts`
- `src/taskpane/welcome-login.ts`
- `src/commands/builtins/overlays.ts`

It works, but it’s harder to refactor and easy to duplicate styling/behavior.

### Proposal
- Introduce a tiny overlay helper (shell + lifecycle):
  - creates backdrop
  - closes on outside click / Escape
  - enforces consistent max-width + padding
- Move most inline styles into theme CSS classes.
- (Optional later) convert overlays to small Lit components once the patterns stabilize.

### Definition of done
- Less HTML string blob code.
- Less inline styling.
- Consistent overlay behavior across commands + welcome screen.

---

## Suggested execution order (small, low-risk commits)

### Phase 0 — eliminate drift (1–2 commits)
- Unify tool count comments.
- Add `comments` to Excel tool renderer/humanizers.
- Remove leftover references to removed tools.

### Phase 1 — capability registry unification (extension-ready)
- Introduce a canonical registry (start with `src/tools/registry.ts`, but design it to accept **extension contributions** later).
- Make UI registration derive from the same registry (renderers + input humanizers).
- Keep room for future tiering/progressive disclosure (#18) without duplicating lists.

### Phase 1.5 — workbook context primitive (foundation for per-workbook sessions)
- Introduce `src/workbook/context.ts` (or similar) that provides a best-effort `{ workbookId, workbookName, workbookUrl? }`.
- Thread this into session metadata in a backwards-compatible way (even if the UI still shows a global list initially).
- Ensure the design supports “local-only” identity by default, with an opt-in workbook-attached ID later if desired.

### Phase 2 — structured tool result metadata
- Add `details` payloads to a few high-value tools first (`write_cells`, `fill_formula`, `format_cells`).
- Update `tool-renderers.ts` to rely on `details` instead of parsing.

### Phase 3 — bundle size + Node leakage
- Replace side-effect imports where possible.
- Add targeted stubs/aliases for browser build stability.

### Phase 4 — excel/ops extraction
- Extract shared range/worksheet helpers.
- Move repeated logic out of tools.

### Phase 5 — overlay UI standardization
- Overlay helper + CSS class consolidation.

---

## Concrete implementation plan (PR-sized slices)

This section translates the phases above into PR-sized work items with clear boundaries.

### PR 1 — Phase 0 + Phase 1: eliminate drift + introduce a capability registry (extension-ready)

**Goal:** fix existing drift bugs and make “what tools exist” a single source of truth.

**Scope (expected files):**
- `src/tools/registry.ts` *(new)*
  - exports `CORE_TOOL_NAMES` (canonical list)
  - exports `CoreToolName` type
  - exports `createCoreTools()` (canonical core tool creation)
- `src/tools/index.ts`
  - becomes a thin adapter around `createCoreTools()`
  - fixes the “10 tools” comment
- `src/ui/tool-renderers.ts`
  - imports `CORE_TOOL_NAMES` (removes local `EXCEL_TOOL_NAMES`)
  - includes `comments` automatically
  - removes stale `get_recent_changes` reference in `describeToolCall()`
- `src/ui/humanize-params.ts`
  - adds a `comments` input humanizer
  - types the humanizer registry as `Record<CoreToolName, …>` so missing tools fail fast at compile time

**Design constraint:** keep the registry **UI-free** (no Lit/renderer types in `src/tools/*`). UI imports the canonical names/type.

**DoD:**
- `comments` tool is rendered + humanized in the UI.
- No remaining “removed tool” references.
- Only one canonical list of core tool names.

### PR 2 — Phase 2: structured tool results via `ToolResultMessage.details` (additive, no text changes)

**Agreement:** Phase 2 should be **additive metadata only** — do not change the human-readable markdown output, only add stable `details` fields.

**Scope:**
- Add minimal `details` payloads to:
  - `write_cells`
  - `fill_formula`
  - `format_cells`
- Update `src/ui/tool-renderers.ts` to prefer `result.details` for:
  - written/fill address
  - blocked state
  - formula error counts
- Keep a fallback path for older persisted sessions that have no `details`.

**DoD:** renderer no longer needs regex parsing for those fields when `details` is present.

### PR 3 — Phase 1.5: workbook context primitive + session/workbook association (foundation)

**Goal:** create a single place to answer “which workbook is this?” and support workbook-scoped UX later.

**Scope (draft):**
- Add `src/workbook/context.ts` (or similar) returning a best-effort:
  - `{ workbookId, workbookName?, workbookUrl? }`
- **Workbook identity (default):** local-only by default.
  - Use `Office.context.document.url` when present, but store a **hash** (never persist raw paths/URLs).
  - If URL is absent, identity is ephemeral.
- **Manual linking (FYI, future):** support a mechanism where the assistant can *suggest* a link, and the user can manually link/unlink sessions ↔ workbook.
  - “Save As” should carry the link by default, but it can be manually overwritten.
- Storage note: `SessionsStore` metadata schema is fixed (from `pi-web-ui`), so session↔workbook mapping likely lives alongside sessions (e.g. `SettingsStore` key prefix `session.workbook.<sessionId>`), not inside metadata.

---

## Verification checklist (every phase)
- `npm run check`
- `npm run build`
- `npm run test:models`
- Manual smoke test in Excel:
  - open taskpane
  - connect provider / enter API key
  - send a message
  - run a couple of tools (read/write/format)
  - model selector still works
  - session resume still works

---

## Open questions / decisions to clarify before deeper work
- Are small behavior changes acceptable if they *only* affect tool output phrasing (not Excel operations)?
- Workbook identity: local-only by default vs writing a non-sensitive ID into the workbook (opt-in). If we write, where (custom property vs hidden sheet vs named item)?
- Artifacts/files: default scoping (global workspace + per-workbook tags), file size limits, and whether binary write is allowed or text-only for v1.
- Extensibility: permission/sandbox model for user-supplied code in hosted builds (full access vs scoped API). **V1** can ship with paste-code only; **V2** should add install-from-URL (GitHub) and possibly package sources (npm).
- Bundling: do we prefer a purely local solution, or are we willing to propose a small upstream `pi-web-ui` change for “minimal registration” entrypoints?
- Do we want `system-prompt.ts` tool list to be generated from the registry, or keep it hand-curated for readability?
