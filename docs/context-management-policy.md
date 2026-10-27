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
- Runtime capability refresh in `src/taskpane/init.ts` now assigns `agent.state.tools = ...` only when tool metadata fingerprint changes **or** extension tool revision changes (extension register/unregister/reload), avoiding no-op churn on unrelated refresh passes while preserving schema-stable hot-reload correctness.
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
| Workbook structural context | Stored `autoContext` note in the transcript (not the base prompt) | First request, workbook switch, blueprint text changed, or the note was compacted away |
| Files workspace summary | Stored `autoContext` note | First request with relevant files, signature changed, or compacted away |
| Edits by the user or a collaborator | Stored `autoContext` note (event); Pi's own writes excluded | Before any request (including mid-turn) when the per-tab cursor has new entries |
| Current selection | Not stored; appended as the last message, bounded (address + size, ≤ 20 sampled rows, ≤ 4,000-char table) | Every request |
| Tool results in model-facing history | Keep fresh full detail short-term, summarize/prune older bulky outputs | On pressure/threshold |
| Compaction | Trigger before hard limits to protect quality | Soft and hard thresholds |

---

## Auto-context contract (2026-10-26)

`src/taskpane/context-injection.ts` owns what Pi tells the agent without being asked. Two layers:

1. **Stored notes** (`prepareRequest`, role `autoContext`, `src/messages/auto-context.ts`). Snapshots (workbook blueprint, Files summary) and events (edits by someone other than Pi) are inserted into the runtime transcript before the user's prompt, or after the latest tool results mid-turn. A snapshot is added only when it differs from the last note of its kind in the transcript, so it is sent once, stays visible for the rest of the turn and later turns, persists with the session, and becomes part of the cached prefix. Later notes are appended, never rewritten, so earlier prefixes stay valid.
2. **Selection** (`transformContext`). Re-read for every request and appended last; never stored. Everything before it is byte-identical to the previous request's messages.

Rules:
- A source that fails or times out is logged (`[auto-context]`) and, for the selection, reported to the agent as unavailable. Never silently dropped.
- Auto-context Excel reads use `excelRun(..., { waitForCellEdit: false })`: they must not hold up a model request while the user edits a cell.
- Edit tracking: one workbook-level `worksheets.onChanged` listener per taskpane; each runtime (chat tab) reads the shared log from its own cursor and is told about every edit it did not make itself. Every write Pi makes is recorded as a write window: tool calls through the coordinator wrapper (with their runtime) and user restores from `/revert` or the Backups dialog. Writes are serialized by the workbook lock, so:
  - `triggerSource: "ThisLocalAddin"` edits are matched to the latest write window covering their time (range overlap breaks ties): the runtime's own writes are left out; restores are listed as "Restored from a backup by the user", other tabs' writes as "By Pi in another chat tab", unmatched add-in edits (extensions) as "By Pi outside this chat".
  - Once a co-author has joined, Pi's writes arrive as `"Unknown"` (and stay so after they leave). A local `"Unknown"` edit is matched the same way only when it also overlaps the range the write recorded (`outputAddress`/`inputAddress` from the change audit, or the restore's address; unknown range = any cell); otherwise it is the user's.
  - `source: "Remote"` edits are listed under "By a co-author"; a co-author's row/column/cell insert or delete adds a warning that addresses may have moved.
  - Structure events from anyone (sheet added/deleted/renamed/moved/shown, row/column insert/delete) invalidate the blueprint cache.
- Agent-facing text; never routed through `t()`.

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
- **Current rollout (v2):**
  - **execution-time guardrail (primary):** global tool-output truncation wrapper on all registered tools with Pi-aligned limits (**50KB UTF-8 bytes** or **2000 lines**, whichever first). For context windows **below 128k**, caps scale linearly with the window (floors: 8KB / 200 lines) — see `src/context/window-budgets.ts` (#566).
  - **history shaping (secondary):** keep latest **6** tool results untouched (scaled down for <128k windows, e.g. **3** at 65k, floor 2); the boundary moves in steps of that count, so between N and 2N−1 stay whole and already-sent results are rewritten only once every N results (prompt cache); compact older tool results when payload exceeds **1,200 chars** or contains images; include a deterministic **500-char preview** in compacted form.
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
- **Current rollout (v2):**
  - hard trigger = `min(contextWindow - reserveTokens, qualityCap)`
  - `qualityCap` = **88%** of context window for ≥128k models, **85%** for ≥200k models
  - soft warning = max(70% of hard trigger, hard trigger − 5% of context window, min margin 2,048 tokens)
  - auto-compaction uses the hard trigger both **before queued prompts** and **mid-turn between tool-loop continuations** (`Agent.prepareNextTurn`); status-bar warnings remain on the existing 40%/60% UX thresholds
  - runs that still end in a provider **context-overflow error** get one compact-and-retry recovery pass (see `src/compaction/overflow-recovery.ts`, #566)
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
| 2) Mid-session model switching | **Implement** cache-safe behavior | ✅ shipped (#428, #442) | Default now matches pi-mono (in-place); optional fork-to-new-tab behavior is available as an advanced setting for non-empty sessions. See `docs/upstream-divergences.md` §1. |
| 3) Mid-session toolset churn | **Implement** targeted stabilization | ✅ shipped (#436), refined (#444) | Runtime skips no-op `setTools(...)` updates via fingerprinting and uses extension tool revision tracking for schema-stable hot-reload updates (without blanket eager refreshes). See `docs/upstream-divergences.md` §2. |
| 4) Mid-session system-prompt churn | **Keep + defer deeper refactor** | ✅ decision recorded | Keep dynamic safety-critical sections (rules, execution mode, connection/integration/skills state) in system prompt for now. Defer stable-base + volatile-message layering until telemetry justifies complexity. |
| 5) Side LLM operations (`llm.complete`) | **Keep intentionally independent** | ✅ guidance + session-key isolation implemented | Treat extension side-completions as separate from main runtime context; extension calls now use extension-scoped side session keys so observability/prefix churn is isolated from the primary runtime. See `docs/upstream-divergences.md` §3. |
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
