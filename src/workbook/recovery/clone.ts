/** Clone helpers for recovery state snapshots. */

import type {
  RecoveryCommentThreadState,
  RecoveryConditionalColorScaleCriterion,
  RecoveryConditionalColorScaleState,
  RecoveryConditionalDataBarRule,
  RecoveryConditionalDataBarState,
  RecoveryConditionalFormatRule,
  RecoveryConditionalIcon,
  RecoveryConditionalIconCriterion,
  RecoveryConditionalIconSetState,
  RecoveryFormatAreaState,
  RecoveryFormatBorderState,
  RecoveryFormatRangeState,
  RecoveryFormatSelection,
  RecoveryModifyStructureState,
  RecoveryStructureValueRangeState,
} from "./types.js";

function cloneRecoveryConditionalDataBarRule(rule: RecoveryConditionalDataBarRule): RecoveryConditionalDataBarRule {
  return {
    type: rule.type,
    formula: rule.formula,
  };
}

function cloneRecoveryConditionalDataBarState(state: RecoveryConditionalDataBarState): RecoveryConditionalDataBarState {
  return {
    axisColor: state.axisColor,
    axisFormat: state.axisFormat,
    barDirection: state.barDirection,
    showDataBarOnly: state.showDataBarOnly,
    lowerBoundRule: cloneRecoveryConditionalDataBarRule(state.lowerBoundRule),
    upperBoundRule: cloneRecoveryConditionalDataBarRule(state.upperBoundRule),
    positiveFillColor: state.positiveFillColor,
    positiveBorderColor: state.positiveBorderColor,
    positiveGradientFill: state.positiveGradientFill,
    negativeFillColor: state.negativeFillColor,
    negativeBorderColor: state.negativeBorderColor,
    negativeMatchPositiveFillColor: state.negativeMatchPositiveFillColor,
    negativeMatchPositiveBorderColor: state.negativeMatchPositiveBorderColor,
  };
}

function cloneRecoveryConditionalColorScaleCriterion(
  criterion: RecoveryConditionalColorScaleCriterion,
): RecoveryConditionalColorScaleCriterion {
  return {
    type: criterion.type,
    formula: criterion.formula,
    color: criterion.color,
  };
}

function cloneRecoveryConditionalColorScaleState(
  state: RecoveryConditionalColorScaleState,
): RecoveryConditionalColorScaleState {
  return {
    minimum: cloneRecoveryConditionalColorScaleCriterion(state.minimum),
    midpoint: state.midpoint ? cloneRecoveryConditionalColorScaleCriterion(state.midpoint) : undefined,
    maximum: cloneRecoveryConditionalColorScaleCriterion(state.maximum),
  };
}

function cloneRecoveryConditionalIcon(icon: RecoveryConditionalIcon): RecoveryConditionalIcon {
  return {
    set: icon.set,
    index: icon.index,
  };
}

function cloneRecoveryConditionalIconCriterion(
  criterion: RecoveryConditionalIconCriterion,
): RecoveryConditionalIconCriterion {
  return {
    type: criterion.type,
    operator: criterion.operator,
    formula: criterion.formula,
    customIcon: criterion.customIcon ? cloneRecoveryConditionalIcon(criterion.customIcon) : undefined,
  };
}

function cloneRecoveryConditionalIconSetState(state: RecoveryConditionalIconSetState): RecoveryConditionalIconSetState {
  return {
    style: state.style,
    reverseIconOrder: state.reverseIconOrder,
    showIconOnly: state.showIconOnly,
    criteria: state.criteria.map((criterion) => cloneRecoveryConditionalIconCriterion(criterion)),
  };
}

function cloneRecoveryConditionalFormatRule(rule: RecoveryConditionalFormatRule): RecoveryConditionalFormatRule {
  return {
    type: rule.type,
    stopIfTrue: rule.stopIfTrue,
    formula: rule.formula,
    operator: rule.operator,
    formula1: rule.formula1,
    formula2: rule.formula2,
    textOperator: rule.textOperator,
    text: rule.text,
    topBottomType: rule.topBottomType,
    rank: rule.rank,
    presetCriterion: rule.presetCriterion,
    dataBar: rule.dataBar ? cloneRecoveryConditionalDataBarState(rule.dataBar) : undefined,
    colorScale: rule.colorScale ? cloneRecoveryConditionalColorScaleState(rule.colorScale) : undefined,
    iconSet: rule.iconSet ? cloneRecoveryConditionalIconSetState(rule.iconSet) : undefined,
    fillColor: rule.fillColor,
    fontColor: rule.fontColor,
    bold: rule.bold,
    italic: rule.italic,
    underline: rule.underline,
    appliesToAddress: rule.appliesToAddress,
  };
}

export function cloneRecoveryConditionalFormatRules(
  rules: readonly RecoveryConditionalFormatRule[],
): RecoveryConditionalFormatRule[] {
  return rules.map((rule) => cloneRecoveryConditionalFormatRule(rule));
}

export function cloneRecoveryCommentThreadState(state: RecoveryCommentThreadState): RecoveryCommentThreadState {
  return {
    exists: state.exists,
    content: state.content,
    resolved: state.resolved,
    replies: [...state.replies],
  };
}

function cloneUnknownGrid(grid: readonly unknown[][]): unknown[][] {
  return grid.map((row) => [...row]);
}

function cloneRecoveryStructureValueRangeState(
  dataRange: RecoveryStructureValueRangeState,
): RecoveryStructureValueRangeState {
  return {
    address: dataRange.address,
    rowCount: dataRange.rowCount,
    columnCount: dataRange.columnCount,
    values: cloneUnknownGrid(dataRange.values),
    formulas: cloneUnknownGrid(dataRange.formulas),
  };
}

export function cloneRecoveryModifyStructureState(state: RecoveryModifyStructureState): RecoveryModifyStructureState {
  switch (state.kind) {
    case "sheet_name":
      return {
        kind: "sheet_name",
        sheetId: state.sheetId,
        name: state.name,
      };
    case "sheet_visibility":
      return {
        kind: "sheet_visibility",
        sheetId: state.sheetId,
        visibility: state.visibility,
      };
    case "sheet_absent":
      return {
        kind: "sheet_absent",
        sheetId: state.sheetId,
        sheetName: state.sheetName,
        ...(state.allowDataDelete === undefined ? {} : { allowDataDelete: state.allowDataDelete }),
      };
    case "sheet_present":
      return {
        kind: "sheet_present",
        sheetId: state.sheetId,
        sheetName: state.sheetName,
        position: state.position,
        visibility: state.visibility,
        ...(state.dataRange ? { dataRange: cloneRecoveryStructureValueRangeState(state.dataRange) } : {}),
      };
    case "rows_absent":
      return {
        kind: "rows_absent",
        sheetId: state.sheetId,
        sheetName: state.sheetName,
        position: state.position,
        count: state.count,
        ...(state.allowDataDelete === undefined ? {} : { allowDataDelete: state.allowDataDelete }),
      };
    case "rows_present":
      return {
        kind: "rows_present",
        sheetId: state.sheetId,
        sheetName: state.sheetName,
        position: state.position,
        count: state.count,
        ...(state.dataRange ? { dataRange: cloneRecoveryStructureValueRangeState(state.dataRange) } : {}),
      };
    case "columns_absent":
      return {
        kind: "columns_absent",
        sheetId: state.sheetId,
        sheetName: state.sheetName,
        position: state.position,
        count: state.count,
        ...(state.allowDataDelete === undefined ? {} : { allowDataDelete: state.allowDataDelete }),
      };
    case "columns_present":
      return {
        kind: "columns_present",
        sheetId: state.sheetId,
        sheetName: state.sheetName,
        position: state.position,
        count: state.count,
        ...(state.dataRange ? { dataRange: cloneRecoveryStructureValueRangeState(state.dataRange) } : {}),
      };
  }
}

export function cloneRecoveryFormatSelection(selection: RecoveryFormatSelection): RecoveryFormatSelection {
  return {
    numberFormat: selection.numberFormat,
    fillColor: selection.fillColor,
    fontColor: selection.fontColor,
    bold: selection.bold,
    italic: selection.italic,
    underlineStyle: selection.underlineStyle,
    fontName: selection.fontName,
    fontSize: selection.fontSize,
    horizontalAlignment: selection.horizontalAlignment,
    verticalAlignment: selection.verticalAlignment,
    wrapText: selection.wrapText,
    columnWidth: selection.columnWidth,
    rowHeight: selection.rowHeight,
    mergedAreas: selection.mergedAreas,
    borderTop: selection.borderTop,
    borderBottom: selection.borderBottom,
    borderLeft: selection.borderLeft,
    borderRight: selection.borderRight,
    borderInsideHorizontal: selection.borderInsideHorizontal,
    borderInsideVertical: selection.borderInsideVertical,
  };
}

function cloneRecoveryFormatBorderState(state: RecoveryFormatBorderState): RecoveryFormatBorderState {
  return {
    style: state.style,
    weight: state.weight,
    color: state.color,
  };
}

export function cloneStringGrid(grid: readonly string[][]): string[][] {
  return grid.map((row) => [...row]);
}

function cloneRecoveryFormatAreaState(area: RecoveryFormatAreaState): RecoveryFormatAreaState {
  return {
    address: area.address,
    rowCount: area.rowCount,
    columnCount: area.columnCount,
    numberFormat: area.numberFormat ? cloneStringGrid(area.numberFormat) : undefined,
    fillColor: area.fillColor,
    fontColor: area.fontColor,
    bold: area.bold,
    italic: area.italic,
    underlineStyle: area.underlineStyle,
    fontName: area.fontName,
    fontSize: area.fontSize,
    horizontalAlignment: area.horizontalAlignment,
    verticalAlignment: area.verticalAlignment,
    wrapText: area.wrapText,
    columnWidths: area.columnWidths ? [...area.columnWidths] : undefined,
    rowHeights: area.rowHeights ? [...area.rowHeights] : undefined,
    mergedAreas: area.mergedAreas ? [...area.mergedAreas] : undefined,
    borderTop: area.borderTop ? cloneRecoveryFormatBorderState(area.borderTop) : undefined,
    borderBottom: area.borderBottom ? cloneRecoveryFormatBorderState(area.borderBottom) : undefined,
    borderLeft: area.borderLeft ? cloneRecoveryFormatBorderState(area.borderLeft) : undefined,
    borderRight: area.borderRight ? cloneRecoveryFormatBorderState(area.borderRight) : undefined,
    borderInsideHorizontal: area.borderInsideHorizontal
      ? cloneRecoveryFormatBorderState(area.borderInsideHorizontal)
      : undefined,
    borderInsideVertical: area.borderInsideVertical
      ? cloneRecoveryFormatBorderState(area.borderInsideVertical)
      : undefined,
  };
}

export function cloneRecoveryFormatRangeState(state: RecoveryFormatRangeState): RecoveryFormatRangeState {
  return {
    selection: cloneRecoveryFormatSelection(state.selection),
    areas: state.areas.map((area) => cloneRecoveryFormatAreaState(area)),
    cellCount: state.cellCount,
  };
}
