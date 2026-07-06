/**
 * Restore strategy helpers for workbook recovery snapshots.
 */

import type { WorkbookContext } from "../context.js";
import type {
  AppendChartRecoverySnapshotArgs,
  AppendCommentThreadRecoverySnapshotArgs,
  AppendConditionalFormatRecoverySnapshotArgs,
  AppendFormatCellsRecoverySnapshotArgs,
  AppendModifyStructureRecoverySnapshotArgs,
  AppendWorkbookRecoverySnapshotArgs,
  RestoreWorkbookRecoverySnapshotResult,
  WorkbookRecoverySnapshot,
  WorkbookRecoverySnapshotKind,
} from "../recovery-log.js";
import type {
  RecoveryChartApplyResult,
  RecoveryChartState,
  RecoveryCommentThreadState,
  RecoveryConditionalFormatCaptureResult,
  RecoveryConditionalFormatRule,
  RecoveryFormatRangeState,
  RecoveryModifyStructureState,
} from "../recovery-states.js";

interface WorkbookRangeState {
  values: DynamicValue[][];
  formulas: DynamicValue[][];
}

interface CountChangedCellsArgs {
  beforeValues: DynamicValue[][];
  beforeFormulas: DynamicValue[][];
  afterValues: DynamicValue[][];
  afterFormulas: DynamicValue[][];
}

export interface RestoreWorkbookRecoverySnapshotDependencies {
  applySnapshot: (address: string, values: DynamicValue[][]) => Promise<WorkbookRangeState>;
  applyFormatCellsSnapshot: (
    address: string,
    state: RecoveryFormatRangeState,
  ) => Promise<RecoveryFormatRangeState>;
  applyModifyStructureSnapshot: (
    address: string,
    state: RecoveryModifyStructureState,
  ) => Promise<RecoveryModifyStructureState>;
  applyConditionalFormatSnapshot: (
    address: string,
    rules: RecoveryConditionalFormatRule[],
  ) => Promise<RecoveryConditionalFormatCaptureResult>;
  applyCommentThreadSnapshot: (
    address: string,
    state: RecoveryCommentThreadState,
  ) => Promise<RecoveryCommentThreadState>;
  applyChartSnapshot: (
    address: string,
    state: RecoveryChartState,
  ) => Promise<RecoveryChartApplyResult>;
  appendRangeSnapshot: (
    args: AppendWorkbookRecoverySnapshotArgs,
    workbookContext: WorkbookContext,
  ) => Promise<WorkbookRecoverySnapshot | null>;
  appendFormatCellsSnapshot: (
    args: AppendFormatCellsRecoverySnapshotArgs,
    workbookContext: WorkbookContext,
  ) => Promise<WorkbookRecoverySnapshot | null>;
  appendModifyStructureSnapshot: (
    args: AppendModifyStructureRecoverySnapshotArgs,
    workbookContext: WorkbookContext,
  ) => Promise<WorkbookRecoverySnapshot | null>;
  appendConditionalFormatSnapshot: (
    args: AppendConditionalFormatRecoverySnapshotArgs,
    workbookContext: WorkbookContext,
  ) => Promise<WorkbookRecoverySnapshot | null>;
  appendCommentThreadSnapshot: (
    args: AppendCommentThreadRecoverySnapshotArgs,
    workbookContext: WorkbookContext,
  ) => Promise<WorkbookRecoverySnapshot | null>;
  appendChartSnapshot: (
    args: AppendChartRecoverySnapshotArgs,
    workbookContext: WorkbookContext,
  ) => Promise<WorkbookRecoverySnapshot | null>;
  toRestoreValues: (values: DynamicValue[][], formulas: DynamicValue[][]) => DynamicValue[][];
  countChangedCells: (args: CountChangedCellsArgs) => number;
}

export interface RestoreWorkbookRecoverySnapshotArgs {
  snapshot: WorkbookRecoverySnapshot;
  workbookContext: WorkbookContext;
  dependencies: RestoreWorkbookRecoverySnapshotDependencies;
}

function resolveSnapshotKind(snapshot: WorkbookRecoverySnapshot): WorkbookRecoverySnapshotKind {
  return snapshot.snapshotKind ?? "range_values";
}

function assertSnapshotWorkbookIdentity(snapshot: WorkbookRecoverySnapshot, workbookContext: WorkbookContext): void {
  if (!snapshot.workbookId) {
    throw new Error("Snapshot is missing workbook identity and cannot be restored safely.");
  }

  if (!workbookContext.workbookId) {
    throw new Error("Current workbook identity is unavailable; cannot safely restore this snapshot.");
  }

  if (snapshot.workbookId !== workbookContext.workbookId) {
    throw new Error("Snapshot belongs to a different workbook.");
  }
}

export async function restoreWorkbookRecoverySnapshot(
  args: RestoreWorkbookRecoverySnapshotArgs,
): Promise<RestoreWorkbookRecoverySnapshotResult> {
  const { snapshot, workbookContext, dependencies } = args;

  assertSnapshotWorkbookIdentity(snapshot, workbookContext);

  const snapshotKind = resolveSnapshotKind(snapshot);

  if (snapshotKind === "format_cells_state") {
    const targetState = snapshot.formatRangeState;
    if (!targetState) {
      throw new Error("Format backup data is missing.");
    }

    const currentState = await dependencies.applyFormatCellsSnapshot(snapshot.address, targetState);
    const inverseSnapshot = await dependencies.appendFormatCellsSnapshot(
      {
        toolName: "restore_snapshot",
        toolCallId: `restore:${snapshot.id}`,
        address: snapshot.address,
        changedCount: snapshot.changedCount,
        formatRangeState: currentState,
        restoredFromSnapshotId: snapshot.id,
      },
      workbookContext,
    );

    return {
      restoredSnapshotId: snapshot.id,
      inverseSnapshotId: inverseSnapshot?.id ?? null,
      address: snapshot.address,
      changedCount: snapshot.changedCount,
    };
  }

  if (snapshotKind === "modify_structure_state") {
    const targetState = snapshot.modifyStructureState;
    if (!targetState) {
      throw new Error("Structure backup data is missing.");
    }

    const currentState = await dependencies.applyModifyStructureSnapshot(snapshot.address, targetState);
    const inverseSnapshot = await dependencies.appendModifyStructureSnapshot(
      {
        toolName: "restore_snapshot",
        toolCallId: `restore:${snapshot.id}`,
        address: snapshot.address,
        changedCount: snapshot.changedCount,
        modifyStructureState: currentState,
        restoredFromSnapshotId: snapshot.id,
      },
      workbookContext,
    );

    return {
      restoredSnapshotId: snapshot.id,
      inverseSnapshotId: inverseSnapshot?.id ?? null,
      address: snapshot.address,
      changedCount: snapshot.changedCount,
    };
  }

  if (snapshotKind === "conditional_format_rules") {
    const rules = snapshot.conditionalFormatRules ?? [];
    const currentState = await dependencies.applyConditionalFormatSnapshot(snapshot.address, rules);

    if (!currentState.supported) {
      throw new Error(currentState.reason ?? "Conditional format backup cannot be restored safely.");
    }

    const inverseSnapshot = await dependencies.appendConditionalFormatSnapshot(
      {
        toolName: "restore_snapshot",
        toolCallId: `restore:${snapshot.id}`,
        address: snapshot.address,
        changedCount: snapshot.changedCount,
        cellCount: snapshot.cellCount,
        conditionalFormatRules: currentState.rules,
        restoredFromSnapshotId: snapshot.id,
      },
      workbookContext,
    );

    return {
      restoredSnapshotId: snapshot.id,
      inverseSnapshotId: inverseSnapshot?.id ?? null,
      address: snapshot.address,
      changedCount: snapshot.changedCount,
    };
  }

  if (snapshotKind === "comment_thread") {
    const targetState = snapshot.commentThreadState;
    if (!targetState) {
      throw new Error("Comment backup data is missing.");
    }

    const currentState = await dependencies.applyCommentThreadSnapshot(snapshot.address, targetState);
    const inverseSnapshot = await dependencies.appendCommentThreadSnapshot(
      {
        toolName: "restore_snapshot",
        toolCallId: `restore:${snapshot.id}`,
        address: snapshot.address,
        changedCount: snapshot.changedCount,
        commentThreadState: currentState,
        restoredFromSnapshotId: snapshot.id,
      },
      workbookContext,
    );

    return {
      restoredSnapshotId: snapshot.id,
      inverseSnapshotId: inverseSnapshot?.id ?? null,
      address: snapshot.address,
      changedCount: snapshot.changedCount,
    };
  }

  if (snapshotKind === "chart_state") {
    const targetState = snapshot.chartState;
    if (!targetState) {
      throw new Error("Chart backup data is missing.");
    }

    const applied = await dependencies.applyChartSnapshot(snapshot.address, targetState);
    const inverseSnapshot = applied.state
      ? await dependencies.appendChartSnapshot(
        {
          toolName: "restore_snapshot",
          toolCallId: `restore:${snapshot.id}`,
          // A restore can rename the chart, so the inverse must be stored at
          // the post-restore identity for the rollback backup to resolve.
          address: applied.address,
          changedCount: snapshot.changedCount,
          chartState: applied.state,
          restoredFromSnapshotId: snapshot.id,
        },
        workbookContext,
      )
      : null;

    return {
      restoredSnapshotId: snapshot.id,
      inverseSnapshotId: inverseSnapshot?.id ?? null,
      address: snapshot.address,
      changedCount: snapshot.changedCount,
    };
  }

  const restoreValues = dependencies.toRestoreValues(snapshot.beforeValues, snapshot.beforeFormulas);
  const currentState = await dependencies.applySnapshot(snapshot.address, restoreValues);

  const inverseChangedCount = dependencies.countChangedCells({
    beforeValues: currentState.values,
    beforeFormulas: currentState.formulas,
    afterValues: snapshot.beforeValues,
    afterFormulas: snapshot.beforeFormulas,
  });

  const inverseSnapshot = await dependencies.appendRangeSnapshot(
    {
      toolName: "restore_snapshot",
      toolCallId: `restore:${snapshot.id}`,
      address: snapshot.address,
      changedCount: inverseChangedCount,
      beforeValues: currentState.values,
      beforeFormulas: currentState.formulas,
      restoredFromSnapshotId: snapshot.id,
    },
    workbookContext,
  );

  return {
    restoredSnapshotId: snapshot.id,
    inverseSnapshotId: inverseSnapshot?.id ?? null,
    address: snapshot.address,
    changedCount: inverseChangedCount,
  };
}
