import {
  recoveryCheckpointCreated,
  recoveryCheckpointUnavailable,
} from "../recovery-metadata.js";
import type { WorkbookRecoverySnapshot } from "../../workbook/recovery-log.js";
import type {
  MutationFinalizeDependencies,
  MutationFinalizeOperation,
  MutationFinalizeRecoveryResult,
  MutationRecoveryStep,
  MutationResultDetails,
} from "./types.js";

export async function finalizeMutationOperation<TDetails extends MutationResultDetails>(
  dependencies: MutationFinalizeDependencies,
  operation: MutationFinalizeOperation<TDetails>,
): Promise<MutationFinalizeRecoveryResult | null> {
  const recovery = operation.recovery;

  try {
    await dependencies.appendAuditEntry(operation.auditEntry);
  } catch {
    const warningTarget = operation.auditWarning ?? recovery;
    if (warningTarget) {
      warningTarget.appendResultNote(
        warningTarget.result,
        "⚠ Workbook audit entry could not be saved.",
      );
    }
  }

  if (!recovery) {
    return null;
  }

  return finalizeMutationRecoveryStep(recovery);
}

export async function finalizeMutationRecoveryStep<TDetails extends MutationResultDetails>(
  recovery: MutationRecoveryStep<TDetails>,
): Promise<MutationFinalizeRecoveryResult> {
  let checkpoint: WorkbookRecoverySnapshot | null;
  try {
    checkpoint = await recovery.appendRecoverySnapshot();
  } catch {
    recovery.result.details.recovery = recoveryCheckpointUnavailable("checkpoint_creation_failed");
    recovery.appendResultNote(
      recovery.result,
      "⚠ Recovery checkpoint could not be saved. The workbook change succeeded, but automatic rollback is unavailable.",
    );

    return {
      checkpointCreated: false,
    };
  }

  if (!checkpoint) {
    recovery.result.details.recovery = recoveryCheckpointUnavailable(recovery.unavailableReason);
    recovery.appendResultNote(recovery.result, recovery.unavailableNote);

    return {
      checkpointCreated: false,
    };
  }

  recovery.result.details.recovery = recoveryCheckpointCreated(checkpoint.id);

  if (recovery.dispatchSnapshotCreated) {
    recovery.dispatchSnapshotCreated(checkpoint);
  }

  return {
    checkpointCreated: true,
    snapshotId: checkpoint.id,
  };
}
