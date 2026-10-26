import {
  backupSkippedNote,
  recoveryCheckpointCreated,
  recoveryCheckpointUnavailable,
} from "../recovery-metadata.js";
import type {
  MutationFinalizeDependencies,
  MutationFinalizeOperation,
  MutationFinalizeRecoveryResult,
  MutationRecoveryStep,
  MutationResultDetails,
} from "./types.js";
import { getErrorMessage } from "../../utils/errors.js";

/** Dispatched on `document` when a backup of a completed change could not be stored. */
export const PI_BACKUP_SAVE_FAILED_EVENT = "pi:backup-save-failed";

function notifyBackupSaveFailed(): void {
  if (typeof document === "undefined") return;
  document.dispatchEvent(new Event(PI_BACKUP_SAVE_FAILED_EVENT));
}

export async function finalizeMutationOperation<TDetails extends MutationResultDetails>(
  dependencies: MutationFinalizeDependencies,
  operation: MutationFinalizeOperation<TDetails>,
): Promise<MutationFinalizeRecoveryResult | null> {
  try {
    await dependencies.appendAuditEntry(operation.auditEntry);
  } catch {
    console.warn("[pi] Failed to append workbook mutation audit entry.");
  }

  const recovery = operation.recovery;
  if (!recovery) {
    return null;
  }

  return finalizeMutationRecoveryStep(recovery);
}

export async function finalizeMutationRecoveryStep<TDetails extends MutationResultDetails>(
  recovery: MutationRecoveryStep<TDetails>,
): Promise<MutationFinalizeRecoveryResult> {
  let checkpoint: Awaited<ReturnType<typeof recovery.appendRecoverySnapshot>>;
  try {
    checkpoint = await recovery.appendRecoverySnapshot();
  } catch (error) {
    // The change is already in the workbook; say plainly that it has no
    // backup instead of failing the call or implying it can be reverted.
    const reason = getErrorMessage(error);
    console.warn(`[pi] ${reason}`);
    recovery.result.details.recovery = recoveryCheckpointUnavailable(reason);
    recovery.appendResultNote(
      recovery.result,
      `⚠️ ${reason} This change cannot be reverted from backups.`,
    );
    notifyBackupSaveFailed();
    return {
      checkpointCreated: false,
    };
  }

  if (!checkpoint) {
    recovery.result.details.recovery = recoveryCheckpointUnavailable(recovery.unavailableReason);
    recovery.appendResultNote(recovery.result, recovery.unavailableNote ?? backupSkippedNote(recovery.unavailableReason));

    return {
      checkpointCreated: false,
    };
  }

  recovery.result.details.recovery = recoveryCheckpointCreated(checkpoint.id);

  return {
    checkpointCreated: true,
    snapshotId: checkpoint.id,
  };
}
