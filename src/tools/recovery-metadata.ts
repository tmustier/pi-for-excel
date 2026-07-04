import type { RecoveryCheckpointDetails } from "./tool-details.js";

export const CHECKPOINTED_TOOL_LABEL = "`write_cells`, `fill_formula`, `python_transform_range`, `format_cells`, `conditional_format`, `comments`, `charts` create/update, and supported `modify_structure` actions";

export const NON_CHECKPOINTED_MUTATION_REASON =
  "This mutation type is not yet covered by workbook backups.";

export const NON_CHECKPOINTED_MUTATION_NOTE =
  `ℹ️ Backup not created. \`workbook_history\` currently tracks ${CHECKPOINTED_TOOL_LABEL}.`;

export const CHECKPOINT_SKIPPED_REASON =
  "Backup capture was skipped (workbook identity unavailable or snapshot data unsupported).";

export const CHECKPOINT_SKIPPED_NOTE =
  `ℹ️ Backup not created for this mutation. \`workbook_history\` tracks ${CHECKPOINTED_TOOL_LABEL} when backup capture succeeds.`;

export function recoveryCheckpointCreated(snapshotId: string): RecoveryCheckpointDetails {
  return {
    status: "checkpoint_created",
    snapshotId,
  };
}

export function recoveryCheckpointUnavailable(reason: string): RecoveryCheckpointDetails {
  return {
    status: "not_available",
    reason,
  };
}
