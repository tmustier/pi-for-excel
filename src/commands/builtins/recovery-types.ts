/**
 * Shared recovery checkpoint types, used by the Backups settings page,
 * recovery filtering helpers, and taskpane wiring.
 */

export type RecoveryCheckpointToolName =
  | "write_cells"
  | "fill_formula"
  | "python_transform_range"
  | "format_cells"
  | "conditional_format"
  | "comments"
  | "charts"
  | "modify_structure"
  | "restore_snapshot";

export interface RecoveryCheckpointSummary {
  id: string;
  at: number;
  toolName: RecoveryCheckpointToolName;
  address: string;
  changedCount: number;
  restoredFromSnapshotId?: string;
}
