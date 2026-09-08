import {
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

export async function finalizeMutationOperation<TDetails extends MutationResultDetails>(
  dependencies: MutationFinalizeDependencies,
  operation: MutationFinalizeOperation<TDetails>,
): Promise<MutationFinalizeRecoveryResult | null> {
  await dependencies.appendAuditEntry(operation.auditEntry);

  const recovery = operation.recovery;
  if (!recovery) {
    return null;
  }

  return finalizeMutationRecoveryStep(recovery);
}

export async function finalizeMutationRecoveryStep<TDetails extends MutationResultDetails>(
  recovery: MutationRecoveryStep<TDetails>,
): Promise<MutationFinalizeRecoveryResult> {
  const checkpoint = await recovery.appendRecoverySnapshot();

  if (!checkpoint) {
    recovery.result.details.recovery = recoveryCheckpointUnavailable(recovery.unavailableReason);
    recovery.appendResultNote(recovery.result, recovery.unavailableNote);

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
