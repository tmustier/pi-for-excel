/**
 * Workbook recovery event contracts.
 */

import type { WorkbookRecoveryToolName } from "./recovery-log.js";

export const PI_WORKBOOK_SNAPSHOT_CREATED_EVENT = "pi:workbook-snapshot-created";

export interface WorkbookSnapshotCreatedDetail {
  snapshotId: string;
  toolName: WorkbookRecoveryToolName;
  address: string;
  changedCount: number;
}

export function dispatchWorkbookSnapshotCreated(detail: WorkbookSnapshotCreatedDetail): void {
  if (typeof document === "undefined") return;

  document.dispatchEvent(new CustomEvent<WorkbookSnapshotCreatedDetail>(
    PI_WORKBOOK_SNAPSHOT_CREATED_EVENT,
    { detail },
  ));
}
