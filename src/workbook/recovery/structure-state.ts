/** Structure-state capture/apply for workbook recovery snapshots. */

export { applyModifyStructureState } from "./structure-apply.js";
export {
  captureModifyStructureState,
  captureSheetValueDataRange,
  captureValueDataRange,
  columnNumberToLetter,
  estimateModifyStructureCellCount,
  hasValueDataInRange,
  hasValueDataInSheet,
  isRecoverySheetVisibility,
  type CaptureModifyStructureStateArgs,
  type StructureValueDataCaptureResult,
} from "./structure-capture.js";
