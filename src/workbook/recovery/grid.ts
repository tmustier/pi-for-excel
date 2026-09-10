/** Shared helpers for workbook recovery value/formula grids. */

export interface RecoveryGridStats {
  rows: number;
  cols: number;
  cellCount: number;
}

export function rowLength(grid: readonly unknown[][], row: number): number {
  const rowValues = grid[row];
  return Array.isArray(rowValues) ? rowValues.length : 0;
}

export function valueAt(grid: readonly unknown[][], row: number, col: number): unknown {
  const rowValues = grid[row];
  if (!Array.isArray(rowValues)) {
    return "";
  }

  return col < rowValues.length ? rowValues[col] : "";
}

export function cloneGrid(grid: readonly unknown[][]): unknown[][] {
  return grid.map((row) => {
    if (!Array.isArray(row)) {
      return [];
    }

    return [...row];
  });
}

export function gridStats(values: readonly unknown[][], formulas: readonly unknown[][]): RecoveryGridStats {
  const rows = Math.max(values.length, formulas.length);
  let cols = 0;

  for (let row = 0; row < rows; row += 1) {
    cols = Math.max(cols, rowLength(values, row), rowLength(formulas, row));
  }

  return {
    rows,
    cols,
    cellCount: rows * cols,
  };
}

export function normalizeFormula(raw: unknown): string | undefined {
  if (typeof raw !== "string") {
    return undefined;
  }

  const trimmed = raw.trim();
  if (!trimmed.startsWith("=")) {
    return undefined;
  }

  return trimmed;
}

export function toRestoreValues(values: readonly unknown[][], formulas: readonly unknown[][]): unknown[][] {
  const { rows, cols } = gridStats(values, formulas);
  const restored: unknown[][] = [];

  for (let row = 0; row < rows; row += 1) {
    const outRow: unknown[] = [];

    for (let col = 0; col < cols; col += 1) {
      const formula = normalizeFormula(valueAt(formulas, row, col));
      outRow.push(formula ?? valueAt(values, row, col));
    }

    restored.push(outRow);
  }

  return restored;
}
