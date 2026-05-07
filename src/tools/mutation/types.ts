import type { AgentToolResult } from "@earendil-works/pi-agent-core";
import type { AppendWorkbookChangeAuditEntryArgs } from "../../audit/workbook-change-audit.js";
import type { WorkbookRecoverySnapshot } from "../../workbook/recovery-log.js";
import type { RecoveryCheckpointDetails } from "../tool-details.js";

export interface MutationResultDetails {
  recovery?: RecoveryCheckpointDetails;
}

export type MutationResultNoteAppender<TDetails extends MutationResultDetails> = (
  result: AgentToolResult<TDetails>,
  note: string,
) => void;

export interface MutationFinalizeDependencies {
  appendAuditEntry: (entry: AppendWorkbookChangeAuditEntryArgs) => Promise<void>;
}

export interface MutationRecoveryStep<TDetails extends MutationResultDetails> {
  result: AgentToolResult<TDetails>;
  appendRecoverySnapshot: () => Promise<WorkbookRecoverySnapshot | null>;
  appendResultNote: MutationResultNoteAppender<TDetails>;
  unavailableReason: string;
  unavailableNote: string;
  dispatchSnapshotCreated?: (snapshot: WorkbookRecoverySnapshot) => void;
}

export interface MutationFinalizeOperation<TDetails extends MutationResultDetails> {
  auditEntry: AppendWorkbookChangeAuditEntryArgs;
  recovery?: MutationRecoveryStep<TDetails>;
}

export interface MutationFinalizeRecoveryResult {
  checkpointCreated: boolean;
  snapshotId?: string;
}
