import type { RecoveryFormatAreaShape, RecoveryFormatSelection } from "./types.js";

export function hasSelectedFormatProperty(selection: RecoveryFormatSelection): boolean {
  return (
    selection.numberFormat === true ||
    selection.fillColor === true ||
    selection.fontColor === true ||
    selection.bold === true ||
    selection.italic === true ||
    selection.underlineStyle === true ||
    selection.fontName === true ||
    selection.fontSize === true ||
    selection.horizontalAlignment === true ||
    selection.verticalAlignment === true ||
    selection.wrapText === true ||
    selection.columnWidth === true ||
    selection.rowHeight === true ||
    selection.mergedAreas === true ||
    selection.borderTop === true ||
    selection.borderBottom === true ||
    selection.borderLeft === true ||
    selection.borderRight === true ||
    selection.borderInsideHorizontal === true ||
    selection.borderInsideVertical === true
  );
}

function hasAreaScalarSelection(selection: RecoveryFormatSelection): boolean {
  return (
    selection.fillColor === true ||
    selection.fontColor === true ||
    selection.bold === true ||
    selection.italic === true ||
    selection.underlineStyle === true ||
    selection.fontName === true ||
    selection.fontSize === true ||
    selection.horizontalAlignment === true ||
    selection.verticalAlignment === true ||
    selection.wrapText === true ||
    selection.borderTop === true ||
    selection.borderBottom === true ||
    selection.borderLeft === true ||
    selection.borderRight === true ||
    selection.borderInsideHorizontal === true ||
    selection.borderInsideVertical === true
  );
}

function estimateMergedAreasUnitCount(area: RecoveryFormatAreaShape): number {
  const cellCount = area.rowCount * area.columnCount;

  if (cellCount <= 1) {
    return cellCount;
  }

  return Math.floor(cellCount / 2);
}

export function estimateFormatCaptureCellCount(
  areas: readonly RecoveryFormatAreaShape[],
  selection: RecoveryFormatSelection,
): number {
  const includeAreaScalarUnits = hasAreaScalarSelection(selection);

  return areas.reduce((count, area) => {
    let areaCount = count;

    if (selection.numberFormat === true) {
      areaCount += area.rowCount * area.columnCount;
    }

    if (selection.columnWidth === true) {
      areaCount += area.columnCount;
    }

    if (selection.rowHeight === true) {
      areaCount += area.rowCount;
    }

    if (selection.mergedAreas === true) {
      areaCount += estimateMergedAreasUnitCount(area);
    }

    if (includeAreaScalarUnits) {
      areaCount += 1;
    }

    return areaCount;
  }, 0);
}
