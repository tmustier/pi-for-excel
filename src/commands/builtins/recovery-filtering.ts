/**
 * Pure filtering, search, and sort helpers for recovery overlay.
 */

import type { RecoveryCheckpointSummary, RecoveryCheckpointToolName } from "./recovery-overlay.js";

// ---------------------------------------------------------------------------
// Tool filter
// ---------------------------------------------------------------------------

export type RecoveryToolFilter = "all" | RecoveryCheckpointToolName;

export interface RecoveryToolFilterOption {
  value: RecoveryToolFilter;
  label: string;
  count: number;
}

const TOOL_FILTER_LABELS: Record<RecoveryCheckpointToolName, string> = {
  write_cells: "Write",
  fill_formula: "Fill formula",
  python_transform_range: "Python transform",
  format_cells: "Format cells",
  conditional_format: "Conditional format",
  comments: "Comments",
  modify_structure: "Modify structure",
  restore_snapshot: "Restore",
};

export function buildToolFilterOptions(
  checkpoints: readonly RecoveryCheckpointSummary[],
): RecoveryToolFilterOption[] {
  const counts = new Map<RecoveryCheckpointToolName, number>();

  for (const checkpoint of checkpoints) {
    counts.set(checkpoint.toolName, (counts.get(checkpoint.toolName) ?? 0) + 1);
  }

  const options: RecoveryToolFilterOption[] = [
    { value: "all", label: "All tools", count: checkpoints.length },
  ];

  for (const [tool, label] of Object.entries(TOOL_FILTER_LABELS)) {
    const count = counts.get(tool as RecoveryCheckpointToolName) ?? 0;
    if (count > 0) {
      options.push({ value: tool as RecoveryToolFilter, label, count });
    }
  }

  return options;
}

// ---------------------------------------------------------------------------
// Sort
// ---------------------------------------------------------------------------

export type RecoverySortOrder = "newest" | "oldest";

function sortCheckpoints(
  checkpoints: RecoveryCheckpointSummary[],
  order: RecoverySortOrder,
): RecoveryCheckpointSummary[] {
  const sorted = [...checkpoints];
  if (order === "oldest") {
    sorted.sort((a, b) => a.at - b.at);
  } else {
    sorted.sort((a, b) => b.at - a.at);
  }
  return sorted;
}

// ---------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------

function matchesSearch(
  checkpoint: RecoveryCheckpointSummary,
  query: string,
): boolean {
  if (query.length === 0) return true;

  const lower = query.toLowerCase();
  return (
    checkpoint.id.toLowerCase().includes(lower) ||
    checkpoint.toolName.toLowerCase().includes(lower) ||
    checkpoint.address.toLowerCase().includes(lower) ||
    (TOOL_FILTER_LABELS[checkpoint.toolName]?.toLowerCase().includes(lower) ?? false)
  );
}

// ---------------------------------------------------------------------------
// Combined pipeline
// ---------------------------------------------------------------------------

export interface RecoveryFilterState {
  search: string;
  toolFilter: RecoveryToolFilter;
  sortOrder: RecoverySortOrder;
}

export const DEFAULT_FILTER_STATE: RecoveryFilterState = {
  search: "",
  toolFilter: "all",
  sortOrder: "newest",
};

export function applyRecoveryFilters(
  checkpoints: readonly RecoveryCheckpointSummary[],
  state: RecoveryFilterState,
): RecoveryCheckpointSummary[] {
  let result = [...checkpoints];

  // Tool filter
  if (state.toolFilter !== "all") {
    result = result.filter((c) => c.toolName === state.toolFilter);
  }

  // Search
  const query = state.search.trim();
  if (query.length > 0) {
    result = result.filter((c) => matchesSearch(c, query));
  }

  // Sort
  return sortCheckpoints(result, state.sortOrder);
}
