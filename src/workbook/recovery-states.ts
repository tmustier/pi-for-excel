import { firstCellAddress } from "./recovery/address.js";
import {
  cloneRecoveryChartState,
  cloneRecoveryCommentThreadState,
  cloneRecoveryConditionalFormatRules,
  cloneRecoveryFormatRangeState,
  cloneRecoveryModifyStructureState,
} from "./recovery/clone.js";

export {
  firstCellAddress,
  cloneRecoveryChartState,
  cloneRecoveryCommentThreadState,
  cloneRecoveryConditionalFormatRules,
  cloneRecoveryFormatRangeState,
  cloneRecoveryModifyStructureState,
};

export { estimateFormatCaptureCellCount } from "./recovery/format-selection.js";
export { applyFormatCellsState, captureFormatCellsState } from "./recovery/format-state.js";
export type { CaptureFormatCellsStateOptions } from "./recovery/format-state.js";
export { applyModifyStructureState, captureModifyStructureState } from "./recovery/structure-state.js";
export { applyConditionalFormatState, captureConditionalFormatState } from "./recovery/conditional-format-state.js";
export { applyCommentThreadState, captureCommentThreadState } from "./recovery/comment-state.js";
export { applyChartState, captureChartPresentState } from "./recovery/chart-state.js";

export type {
  RecoveryChartAbsentState,
  RecoveryChartLegendState,
  RecoveryChartPositionState,
  RecoveryChartPresentState,
  RecoveryChartState,
  RecoveryChartTitleState,
  RecoveryColumnsAbsentState,
  RecoveryColumnsPresentState,
  RecoveryCommentThreadState,
  RecoveryConditionalCellValueOperator,
  RecoveryConditionalColorCriterionType,
  RecoveryConditionalColorScaleCriterion,
  RecoveryConditionalColorScaleState,
  RecoveryConditionalDataBarAxisFormat,
  RecoveryConditionalDataBarDirection,
  RecoveryConditionalDataBarRule,
  RecoveryConditionalDataBarRuleType,
  RecoveryConditionalDataBarState,
  RecoveryConditionalFormatRule,
  RecoveryConditionalFormatRuleOfType,
  RecoveryConditionalFormatRuleType,
  RecoveryConditionalIcon,
  RecoveryConditionalIconCriterion,
  RecoveryConditionalIconCriterionOperator,
  RecoveryConditionalIconCriterionType,
  RecoveryConditionalIconSet,
  RecoveryConditionalIconSetState,
  RecoveryConditionalPresetCriterion,
  RecoveryConditionalTextOperator,
  RecoveryConditionalTopBottomCriterionType,
  RecoveryFormatAreaState,
  RecoveryFormatBorderState,
  RecoveryFormatRangeState,
  RecoveryFormatSelection,
  RecoveryModifyStructureState,
  RecoveryRowsAbsentState,
  RecoveryRowsPresentState,
  RecoverySheetAbsentState,
  RecoverySheetNameState,
  RecoverySheetPresentState,
  RecoverySheetVisibility,
  RecoverySheetVisibilityState,
  RecoveryStructureValueRangeState,
} from "./recovery/schemas.js";

import type { RecoveryChartState, RecoveryConditionalFormatRule, RecoveryFormatRangeState } from "./recovery/schemas.js";

export interface RecoveryConditionalFormatCaptureResult {
  supported: boolean;
  rules: RecoveryConditionalFormatRule[];
  reason?: string;
}

/** Result of applying a chart snapshot during restore. */
export interface RecoveryChartApplyResult {
  /** Inverse state captured before the restore, or null when none applies. */
  state: RecoveryChartState | null;
  /**
   * Address of the chart identity after the restore was applied. A restore can
   * rename the chart, so inverse snapshots must be stored at this address — not
   * the pre-restore one — for the rollback backup to be restorable.
   */
  address: string;
}

export interface RecoveryFormatCaptureResult {
  supported: boolean;
  state?: RecoveryFormatRangeState;
  reason?: string;
}

export interface RecoveryFormatAreaShape {
  rowCount: number;
  columnCount: number;
}
