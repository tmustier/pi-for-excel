# Context Management Policy (cache-safe)

**Status:** Active policy (2026-02-11)  
**Scope:** How Pi for Excel builds model context across normal turns, tool loops, and long sessions **without regressing prompt caching**.

---

## Why this exists

We optimize for **answer quality and context headroom first**, then token/cost.

In practice, quality drops when we repeatedly inject low-signal context (large tool schemas, stale tool outputs, oversized workbook snapshots), even if caching reduces billed tokens.

This policy sets clear guardrails so we can improve context quality while preserving cache performance.

### Critical clarification (cache vs context window)

- Prompt caching helps **cost/latency** (prefill reuse), but does **not** increase available context window.
- Cached tokens still count toward context occupancy for that request.
- Optimization decisions must therefore target **context headroom first**, not billed input tokens alone.

---

## Current baseline (implemented)

- Each model call is built from: `systemPrompt + messages + tools`.
- Tool disclosure is deterministic on every call (including tool-result continuations) in `src/auth/stream-proxy.ts` (`selectToolBundle()`): when tools are present, runtime currently sends the full tool set.
- Runtime capability refresh in `src/taskpane/init.ts` now fingerprints tool metadata and only calls `agent.setTools(...)` when the fingerprint changes (avoids no-op tool churn during refresh passes), while still forcing eager refreshes when extension tools are present so hot-reloaded handlers are applied immediately.
- Session IDs are stable per chat runtime (`agent.sessionId`), which is used by providers for cache continuity.
- Status/debug UI already shows payload composition counters (`systemChars`, `toolSchemaChars`, `messageChars`, call count) plus prefix-churn counters (`prefixChanges`, split by model/system/tools).
- Context window estimation uses provider usage anchored by `calculateContextTokens()` (`input + output + cacheRead + cacheWrite`) in `src/utils/context-tokens.ts`.
- Auto-compaction now uses shared hard budgets (`getCompactionThresholds`) for earlier quality protection while preserving existing status-bar warning semantics.

---

## Request-level mental model

For each LLM call, payload is rebuilt as:

`systemPrompt + tools + messages`

Implications:

- Tool schemas are a **per-request fixed overhead** (not an ever-growing chat-history block).
- If tool use must remain possible in a continuation call, tools must be present on that continuation request.
- Keeping `systemPrompt` + tool bundles stable improves prompt-cache reuse across turns.

---

## Cache-preserving invariants (must hold)

1. **Stable session identity**
   - Keep `agent.sessionId` stable for the lifetime of a session.

2. **Stable base prompt inside a session**
   - Treat the base system prompt as immutable during a session.
   - Avoid per-turn noise in the system prompt.

3. **Deterministic tool schemas**
   - Deterministic order for tools/schemas.
   - No random IDs/timestamps in tool descriptions/schemas.

4. **Dynamic context at the tail**
   - Put volatile data (selection, recent edits, latest tool outputs) near the end of message history.

5. **Discrete context resets only**
   - Compaction should be explicit/discrete, not continuous churn.

---

## Policy by context layer

| Layer | Policy | Reinjection trigger |
|---|---|---|
| Base system prompt | Keep minimal and stable per session | Every call (provider APIs are request-based) |
| Tool schemas | Include a deterministic tool set on every call so continuations can keep using tools (current runtime policy: full set) | Every call |
| Workbook structural context | Inject as separate context block (not baked repeatedly into base prompt) | Session start + workbook hash/version change |
| Per-turn auto-context (selection + recent changes) | Keep bounded and high-signal | Per user turn when non-empty |
| Tool results in model-facing history | Keep fresh full detail short-term, summarize/prune older bulky outputs | On pressure/threshold |
| Compaction | Trigger before hard limits to protect quality | Soft and hard thresholds |

---

## Implementation plan (next slices)

### Slice 1 — Payload snapshots (observability first)

**Goal:** make optimization decisions with real payload evidence.

- Add a small ring buffer of recent request snapshots (debug-only).
- Retention defaults:
  - keep the latest **24 request snapshots**
  - keep latest-context entries for up to **24 sessions**
  - rationale: enough history to inspect multi-step tool loops while keeping taskpane memory bounded
- Capture, per call:
  - call index
  - continuation vs first call
  - tools included yes/no
  - section sizes (system/tool/messages)
  - optional provider payload shape via `onPayload` (redacted)

**Success:** we can compare before/after context composition on real sessions without guesswork.

---

### Slice 2 — Cache-safe tool disclosure

**Goal:** maximize cache reuse and avoid intent-based cache key partitioning.

- Keep tool-bundle metadata centralized (`src/tools/capabilities.ts`) for shared UI/prompt metadata and future opt-in routing.
- Runtime disclosure currently prefers cache continuity over schema minimization: when tools are present, expose `full`.
- Continuations still include tools on every call so multi-step tool loops remain intact.
- **Current rollout (v2):**
  - `none` when no tools are present
  - `full` for both core-only and mixed (core + non-core/extension) toolsets

**Success:** stable prompt-cache patterns with no capability gaps across turns.

---

### Slice 3 — Tool-result history shaping

**Goal:** cut transcript noise from large tool outputs.

- Add model-facing truncation/summarization for older or oversized tool results.
- Keep full raw output in UI/tool cards (no loss of user-visible detail).
- Keep recency window for exact details (latest N tool results untouched).
- **Current rollout (v1):**
  - **execution-time guardrail (primary):** global tool-output truncation wrapper on all registered tools with Pi-aligned limits (**50KB UTF-8 bytes** or **2000 lines**, whichever first)
  - **history shaping (secondary):** keep latest **6** tool results untouched; compact older tool results when payload exceeds **1,200 chars** or contains images; include a deterministic **500-char preview** in compacted form.
  - truncated outputs include stable machine metadata (`details.outputTruncation`) and best-effort full-output persistence under Files workspace `.tool-output/...`.

**Success:** lower message-context growth rate with no UX regression.

---

### Slice 4 — Workbook context invalidation policy

**Goal:** refresh structural workbook context only when necessary.

- Compute workbook context hash/version from structural signals.
- Reinject structural context on hash/version change, workbook switch, or explicit refresh.
- Avoid re-sending large workbook snapshots every turn.
- **Current rollout (v1):** workbook blueprint removed from base system prompt; injected via auto-context only on initial call, workbook switch, blueprint invalidation, or when compaction removed prior workbook context.
- **Lean invalidation policy:** only **structure-impact** tool writes invalidate blueprint context (centralized in `execution-policy` + coordinator wrapper), while value/format/comment/view writes do not.

**Success:** fewer large context swings; better cache reuse.

---

### Slice 5 — Compaction tuning + hygiene UX

**Goal:** protect quality earlier in long threads.

- Tune soft/hard compaction thresholds for earlier quality protection.
- Keep compaction summary compact and action-oriented.
- Add easier “summarize + start fresh” flow for noisy sessions.
- **Current rollout (v1):**
  - hard trigger = `min(contextWindow - reserveTokens, qualityCap)`
  - `qualityCap` = **88%** of context window for ≥128k models, **85%** for ≥200k models
  - soft warning = max(70% of hard trigger, hard trigger − 5% of context window, min margin 2,048 tokens)
  - auto-compaction uses hard trigger; status-bar warnings remain on the existing 40%/60% UX thresholds
  - summarized slices are persisted in a UI-only `archivedMessages` bucket with a “Show earlier messages” card (excluded from model context)

**Success:** fewer degraded late-thread responses.

---

## Verification checklist (each slice)

- `npm run check`
- `npm run build`
- `npm run test:models`
- Manual Excel smoke test (read/write/format flow)
- Real-session payload comparison with debug snapshots:
  - tools included where expected by bundle policy (including continuations)
  - `toolSchemaChars` down (target: meaningful reduction)
  - context occupancy trends healthy (`calculateContextTokens`: input/output/cacheRead/cacheWrite)
  - cache usage remains healthy (`cacheRead`/`cacheWrite` trend not regressing)

---

## Non-goals

- We are **not** replacing provider caching behavior.
- We are **not** changing user-visible tool result text as part of metadata-only slices.
- We are **not** introducing transport-level append semantics in this phase.

---

## #424 investigation updates (current)

| Area | Decision | Status | Notes |
|---|---|---|---|
| 1) Compaction call-shape | **Defer** behavior change | ✅ documented | Keep isolated summarizer request for now. Memo: `docs/archive/issue-424-compaction-call-shape.md`. |
| 2) Mid-session model switching | **Implement** cache-safe behavior | ✅ shipped (#428) | Non-empty sessions now fork into a new tab/model runtime; empty sessions still switch in place. |
| 3) Mid-session toolset churn | **Implement** targeted stabilization | ✅ shipped (#436) | Runtime now skips no-op `setTools(...)` updates via fingerprinting, while keeping extension-tool refreshes eager for handler hot-reloads. |
| 4) Mid-session system-prompt churn | **Keep + defer deeper refactor** | ✅ decision recorded | Keep dynamic safety-critical sections (rules, execution mode, connection/integration/skills state) in system prompt for now. Defer stable-base + volatile-message layering until telemetry justifies complexity. |
| 5) Side LLM operations (`llm.complete`) | **Keep intentionally independent** | ✅ decision recorded | Treat extension side-completions as separate from main runtime context; add explicit extension-author guidance on cache impact (see `docs/extensions.md`). Defer parent-prefix reuse mode. |
| 6) Cache observability policy | **Implement v1 workflow policy** | ✅ policy + baseline matrix documented | Use prefix-churn counters + payload snapshots as release/PR smoke signals for context changes. Baselines: `docs/cache-observability-baselines.md`. |

### Cache observability policy (v1)

For context/tool/prompt changes, treat the following as a required investigation checklist (not hard-fail CI gates yet):

- Baseline expectations by scenario: `docs/cache-observability-baselines.md`
- Run-log template: `docs/release-smoke-runs/templates/context-cache-telemetry-template.md`

1. Enable debug mode and capture a short deterministic session (at least 5 calls including one tool loop).
2. Inspect `prefixChanges` and reason breakdown (`model`, `systemPrompt`, `tools`) from status/debug snapshots.
3. Compare observed `prefixChangeReasons` to the baseline scenario matrix.
4. Investigate unexpected churn when:
   - model changes occur without explicit user/model-selector action,
   - system-prompt changes occur without explicit rules/execution-mode/integration/connection/skills updates,
   - tool-schema changes occur without explicit integration/extension/tool-config updates.
5. Record findings in PR summary when a context-shape change is intentional.

## Open decisions

1. Exact tool bundle definitions + routing heuristics.
2. Tool-result shaping thresholds (size and recency).
3. Workbook hash signal set (what counts as structural change).
4. Whether to tighten or relax v1 soft/hard compaction budgets by provider/model family after more live-session telemetry.
