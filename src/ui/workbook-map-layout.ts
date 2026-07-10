/**
 * Workbook map layout math — pure pixel geometry for the map canvas.
 *
 * Kept DOM-free so the scale/hit-testing rules can be unit tested without a
 * browser or Office host.
 */

import type { CellRect, SheetMapBounds } from "../workbook/map-scan.js";

/** Largest cell size in CSS px (small sheets should not become billboards). */
export const MAP_MAX_CELL_SIZE = 14;

/** Draw cell grid lines only when cells are at least this many CSS px. */
export const MAP_GRID_MIN_CELL_SIZE = 7;

export interface MapLayout {
  /** Canvas CSS width in px. */
  width: number;
  /** Canvas CSS height in px. */
  height: number;
  /** CSS px per cell (may be fractional for large sheets). */
  cellSize: number;
}

export interface PixelRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Fit a rowCount × colCount sheet into a maxWidth × maxHeight pixel budget. */
export function computeMapLayout(
  rowCount: number,
  colCount: number,
  maxWidth: number,
  maxHeight: number,
): MapLayout {
  const rows = Math.max(1, rowCount);
  const cols = Math.max(1, colCount);
  const budgetWidth = Math.max(1, maxWidth);
  const budgetHeight = Math.max(1, maxHeight);

  const cellSize = Math.min(budgetWidth / cols, budgetHeight / rows, MAP_MAX_CELL_SIZE);

  return {
    width: Math.max(1, Math.round(cols * cellSize)),
    height: Math.max(1, Math.round(rows * cellSize)),
    cellSize,
  };
}

/**
 * Convert a cell rect (absolute sheet coordinates) into canvas pixels.
 * Sub-pixel blocks are widened to 1px so lone cells (e.g. a single error)
 * stay visible on heavily downscaled maps.
 */
export function rectToPixels(rect: CellRect, bounds: SheetMapBounds, cellSize: number): PixelRect {
  const x = (rect.startCol - bounds.startCol) * cellSize;
  const y = (rect.startRow - bounds.startRow) * cellSize;
  const w = (rect.endCol - rect.startCol + 1) * cellSize;
  const h = (rect.endRow - rect.startRow + 1) * cellSize;

  return {
    x,
    y,
    w: Math.max(1, w),
    h: Math.max(1, h),
  };
}

/**
 * Map a canvas CSS pixel position back to an absolute 0-based cell position.
 * Returns `null` when the point falls outside the mapped sheet area.
 */
export function pointToCell(
  px: number,
  py: number,
  bounds: SheetMapBounds,
  cellSize: number,
): { row: number; col: number } | null {
  if (cellSize <= 0 || px < 0 || py < 0) return null;

  const colOffset = Math.floor(px / cellSize);
  const rowOffset = Math.floor(py / cellSize);
  if (rowOffset >= bounds.rowCount || colOffset >= bounds.colCount) return null;

  return {
    row: bounds.startRow + rowOffset,
    col: bounds.startCol + colOffset,
  };
}
