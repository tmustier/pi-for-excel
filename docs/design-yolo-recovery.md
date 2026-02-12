# Draft: YOLO workflow + workbook recovery (issue #27)

Issue: https://github.com/tmustier/pi-for-excel/issues/27

## Goal

Replace cumbersome up-front approval selectors with a low-friction workflow:

- let the agent move quickly
- make mistakes cheap to recover
- keep rollback user-controlled and auditable

## Alternatives considered

### 1) Pre-execution approval selector (safe mode)
- **Pros:** explicit consent before mutation
- **Cons:** high interaction cost for multi-step edits; interrupts flow; hard to keep concise in narrow sidebar
- **Decision:** not primary UX for now

### 2) Full-file snapshots / Save As each step
- **Pros:** strongest recovery semantics
- **Cons:** expensive; potentially heavy/slow; awkward lifecycle (storage, naming, cleanup)
- **Feasibility note:** Office.js can expose document/file APIs depending on host/runtime, but “checkpoint every tool call as full workbook copy” is not practical as a first slice
- **Decision:** keep as future exploration, not first implementation

### 3) Range-level pre-write checkpoints (selected)
- **Pros:** cheap, deterministic, aligns with tool-level mutations
- **Cons:** initial scope only covers tools that write a contiguous range
- **Decision:** implement now

### 4) Operation log only (diff/audit without restore)
- **Pros:** transparency and exportable history
- **Cons:** does not solve “undo mistake now” by itself
- **Decision:** keep (issue #28), but pair with restorable checkpoints

## Implemented slice

- Automatic checkpoints for successful:
  - `write_cells`
  - `fill_formula`
  - `python_transform_range`
- New tool: `workbook_history`
  - `list`
  - `restore`
  - `delete`
  - `clear`
- One-click UI affordance:
  - post-write action toast with **Revert**
- Restore is itself reversible:
  - restoring creates an inverse `restore_snapshot` checkpoint
- Local persistence:
  - `workbook.recovery-snapshots.v1`
- Safety cap:
  - snapshots are skipped above `MAX_RECOVERY_CELLS` to avoid oversized local state
- Identity guard:
  - checkpoints are only persisted when workbook identity is available; restores reject legacy identity-less checkpoints

## Why this is better than approval selectors for now

- User pays cost **only when needed** (on mistakes), not before every edit.
- Works well with multi-step agent plans and rapid iterations.
- Recovery remains explicit and inspectable through tool output + checkpoint history.

## Follow-ups

1. Extend checkpointing beyond contiguous cell writes (format/structure/comment mutations).
2. Add a dedicated checkpoint history panel (browse/restore/delete/export).
3. Evaluate host-specific full-file snapshot feasibility for coarse-grained restore points.
4. Potentially expose “YOLO mode” toggle once we have both lightweight and strict workflows fully defined.
