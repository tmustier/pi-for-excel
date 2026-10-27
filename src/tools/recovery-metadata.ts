import { MAX_RECOVERY_CELLS } from "../workbook/recovery/constants.js";
import type { RecoveryCheckpointDetails } from "./tool-details.js";

export interface RecoveryAttribution {
  getTurnId?: (toolCallId: string) => string | undefined;
}

export function recoveryTurn(attribution: RecoveryAttribution, toolCallId: string): { turnId?: string } {
  const turnId = attribution.getTurnId?.(toolCallId);
  return turnId === undefined ? {} : { turnId };
}

export const CHECKPOINTED_TOOL_LABEL = "`write_cells`, `fill_formula`, `python_transform_range`, `format_cells`, `conditional_format`, `comments`, `charts` create/update, and supported `modify_structure` actions";

export const NON_CHECKPOINTED_MUTATION_REASON =
  "This mutation type is not yet covered by workbook backups.";

export const NON_CHECKPOINTED_MUTATION_NOTE =
  `ℹ️ Backup not created. \`workbook_history\` currently tracks ${CHECKPOINTED_TOOL_LABEL}.`;

export const CHECKPOINT_SKIPPED_REASON =
  "Pi could not identify this workbook or could not save the backup.";

/** Reason for a change too large for a cell backup; names the limit and the user's alternative. */
export function tooLargeToBackUpReason(cellCount: number): string {
  return `Too large to back up: ${cellCount.toLocaleString("en-US")} cells exceed the ` +
    `${MAX_RECOVERY_CELLS.toLocaleString("en-US")}-cell backup limit. ` +
    "For changes this size the user can run `/backup` first to save a full copy.";
}

/** Result note when no backup was saved: says that `workbook_history` cannot restore the change, and why. */
export function backupSkippedNote(reason: string): string {
  return `ℹ️ No backup was saved, so \`workbook_history\` cannot restore this change. ${reason}`;
}

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
