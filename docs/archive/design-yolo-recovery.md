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
- **Feasibility findings (2026-02):**
  - `getFileAsync(Compressed)` is not uniformly available across hosts (Excel on web is PDF-only for this API surface)
  - Office.js does not expose a simple atomic in-place "replace current workbook from snapshot" API
  - per-mutation full-file capture would be operationally heavy (slice IO + storage overhead)
- **Decision:** do not use per-mutation full-file snapshots as baseline; if needed later, consider optional/manual desktop-oriented export + open-new-workbook flow

### 3) Range-level pre-write checkpoints (selected)
- **Pros:** cheap, deterministic, aligns with tool-level mutations
- **Cons:** initial scope only covers tools that write a contiguous range
- **Decision:** implement now

### 4) Operation log only (diff/audit without restore)
- **Pros:** transparency and exportable history
- **Cons:** does not solve “undo mistake now” by itself
- **Decision:** keep (issue #28), but pair with restorable checkpoints

## Implemented slice

- Automatic checkpoints for successful mutations:
  - `write_cells`
  - `fill_formula`
  - `python_transform_range`
  - `format_cells`
  - `conditional_format`
  - `comments` (mutating actions)
  - `modify_structure` (all actions, including value-preserving checkpoints for destructive deletes)
- New tool: `workbook_history`
  - `list`
  - `restore`
  - `delete`
  - `clear`
- UI affordances:
  - checkpoint browser overlay (menu + `/history`) for restore/delete/clear
  - `/revert` command for latest-backup rollback
- Restore is itself reversible:
  - restoring creates an inverse `restore_snapshot` checkpoint
- Local persistence:
  - `workbook.recovery-snapshots.v1`
- Safety cap:
  - snapshots are skipped above `MAX_RECOVERY_CELLS` to avoid oversized local state
- Coverage signaling:
  - unsupported mutation tools/actions explicitly state when no checkpoint is created
  - `format_cells` checkpoint coverage now includes merge/unmerge state
  - `conditional_format` checkpoint coverage includes `custom`, `cell_value`, `contains_text`, `top_bottom`, and `preset_criteria` rules
  - `modify_structure` now checkpoints all actions, including destructive deletes (`delete_rows`, `delete_columns`, `delete_sheet`) with captured value/formula payloads when within recovery-size limits
  - structure-absence restores (`sheet_absent`, `rows_absent`, `columns_absent`) are safety-gated and blocked if target data is present, except restore-generated inverse checkpoints that explicitly allow safe value-preserving deletes

## Why this is better than approval selectors for now

- User pays cost **only when needed** (on mistakes), not before every edit.
- Works well with multi-step agent plans and rapid iterations.
- Recovery remains explicit and inspectable through tool output + checkpoint history.

## Follow-ups

1. Enrich checkpoint history UX (search/filter/export, retention controls).
2. If needed, design an optional/manual desktop-oriented full-backup flow (not per-mutation).
3. Potentially expose “YOLO mode” toggle once we have both lightweight and strict workflows fully defined.
