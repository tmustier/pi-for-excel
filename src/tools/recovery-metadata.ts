import type { RecoveryCheckpointDetails } from "./tool-details.js";

export const CHECKPOINTED_TOOL_LABEL = "`write_cells`, `fill_formula`, `python_transform_range`, `conditional_format`, and `comments`";

export const NON_CHECKPOINTED_MUTATION_REASON =
  "This mutation type is not yet covered by workbook checkpoints.";

export const NON_CHECKPOINTED_MUTATION_NOTE =
  `ℹ️ Recovery checkpoint not created. \`workbook_history\` currently tracks ${CHECKPOINTED_TOOL_LABEL}.`;

export const CHECKPOINT_SKIPPED_REASON =
  "Checkpoint capture was skipped (workbook identity unavailable or snapshot data unsupported).";

export const CHECKPOINT_SKIPPED_NOTE =
  `ℹ️ Recovery checkpoint not created for this mutation. \`workbook_history\` tracks ${CHECKPOINTED_TOOL_LABEL} when checkpoint capture succeeds.`;

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
