import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildSheetMapScan,
  emptyRawGroups,
  rectCellCount,
  rectFromAddress,
  rectsFromAddressList,
  splitAddressList,
  sumRectCells,
} from "../src/workbook/map-scan.ts";
import {
  MAP_MAX_CELL_SIZE,
  computeMapLayout,
  pointToCell,
  rectToPixels,
} from "../src/ui/workbook-map-layout.ts";

// ── Address-list splitting ──────────────────────────────────────────────────

void test("splitAddressList splits top-level commas", () => {
  assert.deepEqual(splitAddressList("Sheet1!A1:B3, Sheet1!D5"), ["Sheet1!A1:B3", "Sheet1!D5"]);
});

void test("splitAddressList preserves commas inside quoted sheet names", () => {
  assert.deepEqual(splitAddressList("'P&L, FY26'!A1:B2,Sheet2!C3"), ["'P&L, FY26'!A1:B2", "Sheet2!C3"]);
});

void test("splitAddressList handles doubled quotes in sheet names", () => {
  assert.deepEqual(splitAddressList("'It''s, tricky'!A1,B2"), ["'It''s, tricky'!A1", "B2"]);
});

void test("splitAddressList drops empty segments", () => {
  assert.deepEqual(splitAddressList(""), []);
  assert.deepEqual(splitAddressList("A1,,B2"), ["A1", "B2"]);
});

// ── Rect parsing ────────────────────────────────────────────────────────────

void test("rectFromAddress parses a single cell", () => {
  assert.deepEqual(rectFromAddress("B3"), { startRow: 2, startCol: 1, endRow: 2, endCol: 1 });
});

void test("rectFromAddress parses a range with sheet prefix", () => {
  assert.deepEqual(rectFromAddress("Sheet1!A1:C4"), { startRow: 0, startCol: 0, endRow: 3, endCol: 2 });
});

void test("rectFromAddress parses absolute refs and quoted sheet names", () => {
  assert.deepEqual(rectFromAddress("'P&L, FY26'!$B$2:$D$10"), {
    startRow: 1,
    startCol: 1,
    endRow: 9,
    endCol: 3,
  });
});

void test("rectFromAddress normalizes inverted ranges", () => {
  assert.deepEqual(rectFromAddress("D10:B2"), { startRow: 1, startCol: 1, endRow: 9, endCol: 3 });
});

void test("rectFromAddress rejects full-row/column and malformed refs", () => {
  assert.equal(rectFromAddress("A:B"), null);
  assert.equal(rectFromAddress("1:3"), null);
  assert.equal(rectFromAddress("nonsense"), null);
  assert.equal(rectFromAddress("A1:B2:C3"), null);
});

void test("rectsFromAddressList skips unpaintable segments", () => {
  const rects = rectsFromAddressList("A1:B2,A:A,C3");
  assert.deepEqual(rects, [
    { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
    { startRow: 2, startCol: 2, endRow: 2, endCol: 2 },
  ]);
});

// ── Counting ────────────────────────────────────────────────────────────────

void test("rectCellCount and sumRectCells count inclusive rectangles", () => {
  assert.equal(rectCellCount({ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }), 1);
  assert.equal(rectCellCount({ startRow: 1, startCol: 1, endRow: 3, endCol: 4 }), 12);
  assert.equal(
    sumRectCells([
      { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      { startRow: 1, startCol: 1, endRow: 3, endCol: 4 },
    ]),
    13,
  );
});

void test("buildSheetMapScan merges classes and excludes formula errors from formula count", () => {
  const bounds = { startRow: 0, startCol: 0, rowCount: 10, colCount: 10 };
  const scan = buildSheetMapScan("Model", bounds, {
    formulasAll: [{ startRow: 0, startCol: 0, endRow: 2, endCol: 2 }], // 9 cells
    formulaErrors: [{ startRow: 1, startCol: 1, endRow: 1, endCol: 1 }], // 1 cell
    constantNumbers: [{ startRow: 4, startCol: 0, endRow: 4, endCol: 3 }], // 4 cells
    constantText: [{ startRow: 6, startCol: 0, endRow: 6, endCol: 9 }], // 10 cells
    constantErrors: [{ startRow: 8, startCol: 8, endRow: 8, endCol: 8 }], // 1 cell
  });

  assert.equal(scan.sheetName, "Model");
  assert.equal(scan.counts.formula, 8);
  assert.equal(scan.counts.input, 4);
  assert.equal(scan.counts.label, 10);
  assert.equal(scan.counts.error, 2);
  assert.equal(scan.counts.total, 100);
  assert.equal(scan.counts.empty, 100 - 8 - 4 - 10 - 2);
  assert.equal(scan.rects.error.length, 2);
});

void test("buildSheetMapScan handles empty sheets", () => {
  const scan = buildSheetMapScan("Blank", null, emptyRawGroups());
  assert.equal(scan.bounds, null);
  assert.deepEqual(scan.counts, { formula: 0, input: 0, label: 0, error: 0, empty: 0, total: 0 });
});

// ── Layout geometry ─────────────────────────────────────────────────────────

void test("computeMapLayout fits within the pixel budget", () => {
  const layout = computeMapLayout(1000, 50, 300, 340);
  assert.ok(layout.width <= 300);
  assert.ok(layout.height <= 340);
  assert.equal(layout.cellSize, 340 / 1000);
});

void test("computeMapLayout caps cell size for tiny sheets", () => {
  const layout = computeMapLayout(3, 3, 300, 340);
  assert.equal(layout.cellSize, MAP_MAX_CELL_SIZE);
  assert.equal(layout.width, 3 * MAP_MAX_CELL_SIZE);
  assert.equal(layout.height, 3 * MAP_MAX_CELL_SIZE);
});

void test("computeMapLayout never returns zero-sized canvases", () => {
  const layout = computeMapLayout(1, 1, 1, 1);
  assert.ok(layout.width >= 1);
  assert.ok(layout.height >= 1);
});

void test("rectToPixels offsets by used-range origin and keeps 1px minimum", () => {
  const bounds = { startRow: 5, startCol: 2, rowCount: 100, colCount: 100 };
  const px = rectToPixels({ startRow: 10, startCol: 4, endRow: 14, endCol: 4 }, bounds, 4);
  assert.deepEqual(px, { x: 8, y: 20, w: 4, h: 20 });

  const tiny = rectToPixels({ startRow: 10, startCol: 4, endRow: 10, endCol: 4 }, bounds, 0.1);
  assert.equal(tiny.w, 1);
  assert.equal(tiny.h, 1);
});

void test("pointToCell hit-tests back to absolute sheet coordinates", () => {
  const bounds = { startRow: 5, startCol: 2, rowCount: 10, colCount: 10 };
  assert.deepEqual(pointToCell(0, 0, bounds, 4), { row: 5, col: 2 });
  assert.deepEqual(pointToCell(9, 13, bounds, 4), { row: 8, col: 4 });
  assert.equal(pointToCell(41, 0, bounds, 4), null);
  assert.equal(pointToCell(0, 41, bounds, 4), null);
  assert.equal(pointToCell(-1, 0, bounds, 4), null);
  assert.equal(pointToCell(0, 0, bounds, 0), null);
});

void test("pointToCell round-trips rectToPixels placement", () => {
  const bounds = { startRow: 0, startCol: 0, rowCount: 50, colCount: 20 };
  const cellSize = 6;
  const rect = { startRow: 12, startCol: 7, endRow: 12, endCol: 7 };
  const px = rectToPixels(rect, bounds, cellSize);
  const hit = pointToCell(px.x + 1, px.y + 1, bounds, cellSize);
  assert.deepEqual(hit, { row: 12, col: 7 });
});
