# Design: Multi-Session Workflows in One Workbook (Tabs + Delegate + Team)

> **Status:** Draft
> **Last updated:** 2026-02-10

## Overview

This document specifies how Pi for Excel should support **multiple concurrent agent sessions in the same workbook** from both product and technical standpoints.

Target user outcomes:

1. **Multiple separate sessions** in one workbook (split-tab equivalent)
2. **Background delegate jobs** (subagent-like behavior while the user continues work)
3. **Team workflow** (manager + workers)

---

## Current State (Relevant Baseline)

- One `Agent` instance per taskpane (`src/taskpane/init.ts`)
- One ordered action queue (`src/taskpane/action-queue.ts`)
- Session persistence and workbook association already exist (`src/taskpane/sessions.ts`, `src/workbook/context.ts`, `src/workbook/session-association.ts`)
- Extensions support commands, widgets/overlays, agent event subscriptions, and custom tool registration (`src/commands/extension-api.ts`)
- All Excel tool calls execute directly via `excelRun()` / `Excel.run()` with no cross-session coordinator (`src/excel/helpers.ts`)

Implication: we need a first-class runtime architecture for multi-session concurrency; this cannot be solved with UI-only changes.

---

## Core Invariant

**Parallel reasoning, serialized workbook mutation.**

- Multiple agents may think, plan, and call LLMs concurrently.
- Workbook writes must be coordinated to avoid stale-read/overwrite races.
- Reads may be allowed concurrently in later phases, but write serialization is mandatory.

---

## Goals

1. Allow multiple session runtimes attached to one workbook in one taskpane instance.
2. Preserve predictable workbook behavior under concurrent agent activity.
3. Provide explicit user visibility for “who is editing what” and lock/wait states.
4. Support background delegate execution while main session remains interactive.
5. Provide a team orchestration mode without requiring Node subprocesses.

## Non-Goals (for initial rollout)

- Cross-process locking across separate Excel processes/machines
- Perfect transactional rollback for all Excel operations
- Full distributed conflict-free merge semantics
- Dynamic third-party extension loading for this feature set

---

## User-Facing Modes

### Mode A: Session Tabs (split-tab equivalent)

- User can open multiple chat sessions for the same workbook.
- Each tab has isolated message history and model context.
- Each tab can optionally set a **working scope** (sheet/range pin).
- If a tab attempts a write while another write is in progress, it shows a waiting state.

### Mode B: Delegate Jobs (subagent equivalent)

- Main session can start background jobs that continue while user keeps chatting.
- Jobs have lifecycle states: `queued`, `running`, `waiting_for_lock`, `completed`, `failed`, `cancelled`.
- Initial policy: delegate jobs are read-only or “propose changes”; apply is explicit.

### Mode C: Team Runs

- One manager session coordinates multiple worker sessions.
- Workers gather facts/proposals; manager synthesizes and applies.
- Writes are centralized (manager apply step) in initial team design.

---

## Technical Architecture

### 1) WorkbookCoordinator (new core runtime primitive)

Single coordination point for workbook operations:

- Keyed by `workbookId` (from `getWorkbookContext()`)
- Maintains a write queue/mutex
- Emits operation lifecycle events for UI
- Optionally tracks a monotonic workbook `revision`

Suggested interface:

```ts
export interface WorkbookOperationContext {
  workbookId: string;
  sessionId: string;
  opId: string;
  expectedRevision?: number;
}

export interface WorkbookCoordinator {
  runRead<T>(ctx: WorkbookOperationContext, fn: () => Promise<T>): Promise<T>;
  runWrite<T>(ctx: WorkbookOperationContext, fn: () => Promise<T>): Promise<{ result: T; revision: number }>;
  getRevision(workbookId: string): number;
}
```

Initial policy:
- `runWrite`: strictly serialized
- `runRead`: can execute immediately in Phase 1, or also queued if Office API stability requires it

### 2) Tool Execution Policy (read vs mutate classification)

Every tool call must declare execution class:

- `read`: `get_workbook_overview`, `read_range`, `search_workbook`, `trace_dependencies`, `view_settings` (get mode), `comments` (read mode)
- `mutate`: `write_cells`, `fill_formula`, `modify_structure`, `format_cells`, `conditional_format`, `view_settings` (set mode), `comments` (write/update/reply/delete/resolve)

Mutating tools must execute under `WorkbookCoordinator.runWrite(...)`.

### 3) SessionRuntimeManager (multi-agent in one taskpane)

Owns multiple independent agent runtimes:

```ts
export interface SessionRuntime {
  sessionId: string;
  agent: Agent;
  actionQueue: ReturnType<typeof createActionQueue>;
  scope?: { sheet?: string; range?: string };
}
```

Responsibilities:
- create/remove/switch active session tabs
- isolate prompt queues per session
- route UI events to active session
- keep session persistence integration (`setupSessionPersistence`) per runtime

### 4) DelegateJobManager (background execution)

Creates and tracks non-active session runtimes as jobs.

```ts
export type DelegateJobState =
  | "queued"
  | "running"
  | "waiting_for_lock"
  | "completed"
  | "failed"
  | "cancelled";
```

Responsibilities:
- spawn runtime with task + policy (read-only vs propose)
- surface progress and final summary in main session
- cancellation + timeout support

### 5) TeamOrchestrator

Coordinates manager + worker runtimes:
- split task
- run workers in parallel
- merge worker outputs
- produce manager summary + apply plan

Initial team safety rule:
- Workers do not directly mutate workbook
- Manager performs final apply step through coordinated writes

---

## Data Model and Persistence

### Existing foundation to reuse

- Workbook identity hash (`src/workbook/context.ts`)
- Session↔workbook mapping (`src/workbook/session-association.ts`)

### Additions

1. Active tab state per workbook (for restoring UI layout)
2. Optional saved tab set for each workbook
3. Delegate job history (ephemeral in Phase 1, persistent optional in Phase 2)

Versioned key pattern in `SettingsStore` should be used (e.g. `*.v1.*`).

---

## UX Requirements

1. Show active workbook name in header/status.
2. Show active session tab and busy state.
3. Show coordinator wait state (“Waiting for workbook lock…”).
4. Show source session for each write in debug/audit display.
5. Resume/session picker should default-filter to current workbook.

---

## Rollout Plan

### Phase 1 — Multi-session tabs + write coordinator (MVP)

- Add `SessionRuntimeManager`
- Add `WorkbookCoordinator` with serialized writes
- Add tab UI + per-session queueing
- Keep one visible active session; no background delegates yet

#### Phase 1 execution plan (file-by-file)

##### Workstream A — Workbook coordinator + tool execution policy

**New files**
- `src/workbook/coordinator.ts`
  - Implement write mutex/queue keyed by `workbookId`
  - Expose operation lifecycle events (`queued`, `started`, `completed`, `failed`)
  - Track optional monotonic `revision` per workbook
- `src/tools/execution-policy.ts`
  - Classify tool calls as `read` or `mutate`
  - Handle param-based mode for mixed tools (`comments`, `view_settings`)
- `src/tools/with-workbook-coordinator.ts`
  - Wrap tool `execute()` so mutate calls run under `WorkbookCoordinator.runWrite(...)`

**Touched files**
- `src/taskpane/init.ts`
  - Create coordinator and wrap core tools before agent creation

**Checkpoint A**
- Two concurrent mutating tool calls from separate agents never overlap; second call enters `queued` then runs after first completes.

##### Workstream B — Multi-runtime session manager (tabs backend)

**New files**
- `src/taskpane/session-runtime-manager.ts`
  - Manage runtime map: `{ sessionId, agent, actionQueue, queueDisplay, metadata }`
  - APIs: `createRuntime`, `switchRuntime`, `closeRuntime`, `getActiveRuntime`, `listRuntimes`

**Touched files**
- `src/taskpane/init.ts`
  - Replace single-agent wiring with runtime-manager wiring
- `src/taskpane/action-queue.ts`
  - Keep per-runtime queue semantics; ensure no cross-runtime shared queue state
- `src/taskpane/queue-display.ts`
  - Support runtime-scoped queue display binding (active runtime only)

**Checkpoint B**
- Two runtimes can be created and switched; each preserves independent message history and pending queue state.

##### Workstream C — Active-agent indirection for commands, shortcuts, status

**Touched files**
- `src/commands/builtins/index.ts`
  - Register builtins against an `ActiveAgentProvider` instead of a single captured agent
- `src/commands/builtins/model.ts`
- `src/commands/builtins/clipboard.ts`
- `src/commands/builtins/export.ts`
- `src/commands/builtins/session.ts`
- `src/commands/builtins/overlays.ts`
  - Resume/new behavior targets active runtime manager APIs
- `src/taskpane/keyboard-shortcuts.ts`
  - Resolve active runtime dynamically on each key event
- `src/taskpane/status-bar.ts`
  - Render against current active runtime/agent

**Checkpoint C**
- Switching tabs changes `/model`, `/copy`, `/compact`, `/new`, `/resume`, keyboard shortcuts, and status bar to the newly active runtime.

##### Workstream D — Session tabs UI (MVP)

**New files (optional split)**
- `src/ui/session-tabs.ts` (if extracted)

**Touched files**
- `src/ui/pi-sidebar.ts`
  - Add tab strip UI: create/switch/close
  - Expose callbacks/events for runtime manager
- `src/ui/theme.css` (or existing sidebar CSS module)
  - Tab styles, active state, overflow handling
- `src/taskpane/init.ts`
  - Connect tab events to runtime manager

**Checkpoint D**
- User can create/switch/close tabs; active tab is visually clear; closing active tab activates a deterministic fallback tab.

##### Workstream E — Session persistence + workbook-aware resume in multi-runtime context

**Touched files**
- `src/taskpane/sessions.ts`
  - Support runtime-aware persistence hooks
  - Ensure latest-session mapping updates for the runtime being saved/resumed
- `src/commands/builtins/overlays.ts`
  - Resume dialog defaults to current-workbook sessions with optional “All workbooks” view
- `src/workbook/session-association.ts`
  - Add helper(s) needed for workbook filtering by session id list

**Checkpoint E**
- After reload, latest session for current workbook restores correctly and tab/session association remains intact.

##### Workstream F — Workbook lock UX and telemetry

**Touched files**
- `src/ui/pi-sidebar.ts`
  - Show lock wait message (“Waiting for workbook lock…”) when runtime is queued for mutate op
- `src/taskpane/status-bar.ts`
  - Optional lock indicator/debug info (owner session/op)
- `src/taskpane/init.ts`
  - Wire coordinator events to UI

**Checkpoint F**
- During a long mutate operation in tab A, tab B clearly shows queued lock state for its pending mutate operation.

##### Phase 1 verification checklist

- [ ] Open two tabs; send independent prompts; histories remain isolated.
- [ ] Trigger concurrent writes; writes serialize correctly.
- [ ] `/new` creates a fresh runtime without destroying other tabs.
- [ ] `/resume` can restore into active tab/runtime and preserve workbook association.
- [ ] Status bar and shortcuts target active tab only.
- [ ] No regression in single-tab behavior.

### Phase 2 — Delegate background jobs

- Add `DelegateJobManager`
- Add `/delegate` command and jobs panel
- Start with read-only/propose mode

### Phase 3 — Team orchestration

- Add `/team` command flow
- Manager + worker orchestration
- Manager-controlled apply

---

## Acceptance Criteria

1. User can keep **at least 2 sessions open** in one workbook, each with independent chat context.
2. Concurrent mutating operations are serialized and never interleave unsafely.
3. User can run a background delegate job while continuing in main session.
4. Team flow can run multiple workers and produce one manager-level result.
5. Session/workbook association remains correct after reload and resume.

---

## Risks and Mitigations

1. **Office.js concurrency edge cases**
   - Mitigation: start with conservative serialization and explicit coordinator telemetry.
2. **UX complexity from multiple active threads**
   - Mitigation: Phase 1 tabs only; background/team behind explicit commands.
3. **Context bloat from many workers**
   - Mitigation: worker outputs summarized before insertion into main thread.
4. **Scope confusion (selection-based context)**
   - Mitigation: per-session scope pinning and scope indicator.

---

## Open Questions

1. Should reads also be serialized initially for maximum Office.js stability?
2. Should background jobs be persisted across reloads in V1, or restarted manually?
3. Should worker writes be allowed behind confirmation in Phase 3, or remain manager-only?
4. Do we need workbook revision checks in Phase 1, or only queue-based ordering?

---

## Related Issues

- #31 — multi-workbook semantics + per-workbook chats
- #23 — sessions: history UI + resume per workbook
- #13 — extensions API build-out
- #14 — agent interface/context strategy
- #28 — auditability/diff log (complementary for multi-session trust)