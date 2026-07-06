/** Clone helpers for recovery state snapshots. */

import type {
  RecoveryChartState,
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
  const cloned: RecoveryConditionalDataBarRule = {
    type: rule.type,
  };

  if (rule.formula !== undefined) {
    cloned.formula = rule.formula;
  }

  return cloned;
}

function cloneRecoveryConditionalDataBarState(state: RecoveryConditionalDataBarState): RecoveryConditionalDataBarState {
  const cloned: RecoveryConditionalDataBarState = {
    axisFormat: state.axisFormat,
    barDirection: state.barDirection,
    showDataBarOnly: state.showDataBarOnly,
    lowerBoundRule: cloneRecoveryConditionalDataBarRule(state.lowerBoundRule),
    upperBoundRule: cloneRecoveryConditionalDataBarRule(state.upperBoundRule),
    positiveFillColor: state.positiveFillColor,
    positiveGradientFill: state.positiveGradientFill,
    negativeFillColor: state.negativeFillColor,
    negativeMatchPositiveFillColor: state.negativeMatchPositiveFillColor,
    negativeMatchPositiveBorderColor: state.negativeMatchPositiveBorderColor,
  };

  if (state.axisColor !== undefined) {
    cloned.axisColor = state.axisColor;
  }
  if (state.positiveBorderColor !== undefined) {
    cloned.positiveBorderColor = state.positiveBorderColor;
  }
  if (state.negativeBorderColor !== undefined) {
    cloned.negativeBorderColor = state.negativeBorderColor;
  }

  return cloned;
}

function cloneRecoveryConditionalColorScaleCriterion(
  criterion: RecoveryConditionalColorScaleCriterion,
): RecoveryConditionalColorScaleCriterion {
  const cloned: RecoveryConditionalColorScaleCriterion = {
    type: criterion.type,
  };

  if (criterion.formula !== undefined) {
    cloned.formula = criterion.formula;
  }
  if (criterion.color !== undefined) {
    cloned.color = criterion.color;
  }

  return cloned;
}

function cloneRecoveryConditionalColorScaleState(
  state: RecoveryConditionalColorScaleState,
): RecoveryConditionalColorScaleState {
  const cloned: RecoveryConditionalColorScaleState = {
    minimum: cloneRecoveryConditionalColorScaleCriterion(state.minimum),
    maximum: cloneRecoveryConditionalColorScaleCriterion(state.maximum),
  };

  if (state.midpoint) {
    cloned.midpoint = cloneRecoveryConditionalColorScaleCriterion(state.midpoint);
  }

  return cloned;
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
  const cloned: RecoveryConditionalIconCriterion = {
    type: criterion.type,
    operator: criterion.operator,
    formula: criterion.formula,
  };

  if (criterion.customIcon) {
    cloned.customIcon = cloneRecoveryConditionalIcon(criterion.customIcon);
  }

  return cloned;
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
  const cloned: RecoveryConditionalFormatRule = {
    type: rule.type,
  };

  if (rule.stopIfTrue !== undefined) cloned.stopIfTrue = rule.stopIfTrue;
  if (rule.formula !== undefined) cloned.formula = rule.formula;
  if (rule.operator !== undefined) cloned.operator = rule.operator;
  if (rule.formula1 !== undefined) cloned.formula1 = rule.formula1;
  if (rule.formula2 !== undefined) cloned.formula2 = rule.formula2;
  if (rule.textOperator !== undefined) cloned.textOperator = rule.textOperator;
  if (rule.text !== undefined) cloned.text = rule.text;
  if (rule.topBottomType !== undefined) cloned.topBottomType = rule.topBottomType;
  if (rule.rank !== undefined) cloned.rank = rule.rank;
  if (rule.presetCriterion !== undefined) cloned.presetCriterion = rule.presetCriterion;
  if (rule.dataBar) cloned.dataBar = cloneRecoveryConditionalDataBarState(rule.dataBar);
  if (rule.colorScale) cloned.colorScale = cloneRecoveryConditionalColorScaleState(rule.colorScale);
  if (rule.iconSet) cloned.iconSet = cloneRecoveryConditionalIconSetState(rule.iconSet);
  if (rule.fillColor !== undefined) cloned.fillColor = rule.fillColor;
  if (rule.fontColor !== undefined) cloned.fontColor = rule.fontColor;
  if (rule.bold !== undefined) cloned.bold = rule.bold;
  if (rule.italic !== undefined) cloned.italic = rule.italic;
  if (rule.underline !== undefined) cloned.underline = rule.underline;
  if (rule.appliesToAddress !== undefined) cloned.appliesToAddress = rule.appliesToAddress;

  return cloned;
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

export function cloneRecoveryChartState(state: RecoveryChartState): RecoveryChartState {
  if (state.kind === "chart_absent") {
    const cloned: RecoveryChartState = {
      kind: "chart_absent",
      sheetName: state.sheetName,
      name: state.name,
    };
    if (state.chartId !== undefined) {
      cloned.chartId = state.chartId;
    }
    return cloned;
  }

  const cloned: RecoveryChartState = {
    kind: "chart_present",
    sheetName: state.sheetName,
    name: state.name,
    chartType: state.chartType,
    title: { ...state.title },
    legend: { ...state.legend },
    position: { ...state.position },
  };

  if (state.xAxisTitle) {
    cloned.xAxisTitle = { ...state.xAxisTitle };
  }
  if (state.yAxisTitle) {
    cloned.yAxisTitle = { ...state.yAxisTitle };
  }

  return cloned;
}

function cloneUnknownGrid(grid: readonly DynamicValue[][]): DynamicValue[][] {
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
  const cloned: RecoveryFormatSelection = {};

  if (selection.numberFormat !== undefined) cloned.numberFormat = selection.numberFormat;
  if (selection.fillColor !== undefined) cloned.fillColor = selection.fillColor;
  if (selection.fontColor !== undefined) cloned.fontColor = selection.fontColor;
  if (selection.bold !== undefined) cloned.bold = selection.bold;
  if (selection.italic !== undefined) cloned.italic = selection.italic;
  if (selection.underlineStyle !== undefined) cloned.underlineStyle = selection.underlineStyle;
  if (selection.fontName !== undefined) cloned.fontName = selection.fontName;
  if (selection.fontSize !== undefined) cloned.fontSize = selection.fontSize;
  if (selection.horizontalAlignment !== undefined) cloned.horizontalAlignment = selection.horizontalAlignment;
  if (selection.verticalAlignment !== undefined) cloned.verticalAlignment = selection.verticalAlignment;
  if (selection.wrapText !== undefined) cloned.wrapText = selection.wrapText;
  if (selection.columnWidth !== undefined) cloned.columnWidth = selection.columnWidth;
  if (selection.rowHeight !== undefined) cloned.rowHeight = selection.rowHeight;
  if (selection.mergedAreas !== undefined) cloned.mergedAreas = selection.mergedAreas;
  if (selection.borderTop !== undefined) cloned.borderTop = selection.borderTop;
  if (selection.borderBottom !== undefined) cloned.borderBottom = selection.borderBottom;
  if (selection.borderLeft !== undefined) cloned.borderLeft = selection.borderLeft;
  if (selection.borderRight !== undefined) cloned.borderRight = selection.borderRight;
  if (selection.borderInsideHorizontal !== undefined) cloned.borderInsideHorizontal = selection.borderInsideHorizontal;
  if (selection.borderInsideVertical !== undefined) cloned.borderInsideVertical = selection.borderInsideVertical;

  return cloned;
}

function cloneRecoveryFormatBorderState(state: RecoveryFormatBorderState): RecoveryFormatBorderState {
  const cloned: RecoveryFormatBorderState = {
    style: state.style,
  };

  if (state.weight !== undefined) cloned.weight = state.weight;
  if (state.color !== undefined) cloned.color = state.color;

  return cloned;
}

export function cloneStringGrid(grid: readonly string[][]): string[][] {
  return grid.map((row) => [...row]);
}

function cloneRecoveryFormatAreaState(area: RecoveryFormatAreaState): RecoveryFormatAreaState {
  const cloned: RecoveryFormatAreaState = {
    address: area.address,
    rowCount: area.rowCount,
    columnCount: area.columnCount,
  };

  if (area.numberFormat) cloned.numberFormat = cloneStringGrid(area.numberFormat);
  if (area.fillColor !== undefined) cloned.fillColor = area.fillColor;
  if (area.fontColor !== undefined) cloned.fontColor = area.fontColor;
  if (area.bold !== undefined) cloned.bold = area.bold;
  if (area.italic !== undefined) cloned.italic = area.italic;
  if (area.underlineStyle !== undefined) cloned.underlineStyle = area.underlineStyle;
  if (area.fontName !== undefined) cloned.fontName = area.fontName;
  if (area.fontSize !== undefined) cloned.fontSize = area.fontSize;
  if (area.horizontalAlignment !== undefined) cloned.horizontalAlignment = area.horizontalAlignment;
  if (area.verticalAlignment !== undefined) cloned.verticalAlignment = area.verticalAlignment;
  if (area.wrapText !== undefined) cloned.wrapText = area.wrapText;
  if (area.columnWidths) cloned.columnWidths = [...area.columnWidths];
  if (area.rowHeights) cloned.rowHeights = [...area.rowHeights];
  if (area.mergedAreas) cloned.mergedAreas = [...area.mergedAreas];
  if (area.borderTop) cloned.borderTop = cloneRecoveryFormatBorderState(area.borderTop);
  if (area.borderBottom) cloned.borderBottom = cloneRecoveryFormatBorderState(area.borderBottom);
  if (area.borderLeft) cloned.borderLeft = cloneRecoveryFormatBorderState(area.borderLeft);
  if (area.borderRight) cloned.borderRight = cloneRecoveryFormatBorderState(area.borderRight);
  if (area.borderInsideHorizontal) {
    cloned.borderInsideHorizontal = cloneRecoveryFormatBorderState(area.borderInsideHorizontal);
  }
  if (area.borderInsideVertical) {
    cloned.borderInsideVertical = cloneRecoveryFormatBorderState(area.borderInsideVertical);
  }

  return cloned;
}

export function cloneRecoveryFormatRangeState(state: RecoveryFormatRangeState): RecoveryFormatRangeState {
  return {
    selection: cloneRecoveryFormatSelection(state.selection),
    areas: state.areas.map((area) => cloneRecoveryFormatAreaState(area)),
    cellCount: state.cellCount,
  };
}
