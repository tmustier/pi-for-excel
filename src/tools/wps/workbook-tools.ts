/** WPS JSAPI implementations for the Phase 2 workbook-tool vertical slice. */

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { buildWorkbookCellChangeSummary } from "../../audit/cell-diff.js";
import {
  colToLetter,
  computeRangeAddress,
  padValues,
  parseCell,
  parseRangeRef,
  qualifiedAddress,
} from "../../excel/helpers.js";
import {
  getWpsEtApplication,
  type WpsCountedCollection,
  type WpsEtApplication,
  type WpsEtRange,
  type WpsEtWorkbook,
  type WpsEtWorksheet,
  type WpsRowsOrColumns,
} from "../../host/wps/jsapi.js";
import { getErrorMessage } from "../../utils/errors.js";
import { formatAsMarkdownTable, extractFormulas, findErrors } from "../../utils/format.js";
import type { ReadRangeCsvDetails, WriteCellsDetails } from "../tool-details.js";
import { countOccupiedCells, validateFormula } from "../write-cells.js";

const WPS_NO_BACKUP_NOTICE =
  "ℹ️ WPS workbook backups are not implemented yet; no automatic recovery snapshot was created.";
const WPS_NO_BACKUP_REASON = "WPS workbook backups are not implemented.";

interface ReadRangeParams {
  range: string;
  mode?: "compact" | "csv" | "detailed";
}

interface WorkbookOverviewParams {
  sheet?: string;
}

interface WriteCellsParams {
  start_cell: string;
  values: DynamicValue[][];
  allow_overwrite?: boolean;
}

interface WpsRangeSnapshot {
  sheetName: string;
  address: string;
  rows: number;
  cols: number;
  values: DynamicValue[][];
  formulas: DynamicValue[][];
  numberFormats: DynamicValue[][];
  metadataWarnings: string[];
}

interface InvalidFormula {
  address: string;
  formula: string;
  reason: string;
}

type WriteCellsResult =
  | {
    blocked: true;
    sheetName: string;
    address: string;
    existingCount: number;
    existingValues: DynamicValue[][];
  }
  | {
    blocked: false;
    sheetName: string;
    address: string;
    beforeValues: DynamicValue[][];
    beforeFormulas: DynamicValue[][];
    readBackValues: DynamicValue[][];
    readBackFormulas: DynamicValue[][];
  };

type BlockedWriteCellsResult = Extract<WriteCellsResult, { blocked: true }>;
type SuccessWriteCellsResult = Extract<WriteCellsResult, { blocked: false }>;

function isUnknownArray(value: DynamicValue): value is DynamicValue[] {
  return Array.isArray(value);
}

function isUnknownGrid(value: DynamicValue): value is DynamicValue[][] {
  return isUnknownArray(value) && value.every((row) => Array.isArray(row));
}

function asString(value: DynamicValue): string | null {
  return typeof value === "string" && value.trim().length > 0 ? value : null;
}

function asWorksheet(value: DynamicValue): WpsEtWorksheet | null {
  if (typeof value !== "object" || value === null) return null;
  return value;
}

function asRange(value: DynamicValue): WpsEtRange | null {
  if (typeof value !== "object" || value === null) return null;
  return value;
}

function worksheetName(sheet: WpsEtWorksheet): string {
  return asString(sheet.Name) ?? asString(sheet.name) ?? "(unnamed sheet)";
}

function workbookName(workbook: WpsEtWorkbook): string {
  return asString(workbook.Name) ?? asString(workbook.name) ?? "(unnamed workbook)";
}

function getCount(collection: WpsCountedCollection | WpsRowsOrColumns | null | undefined): number | null {
  const raw = collection?.Count ?? collection?.count;
  return typeof raw === "number" && Number.isFinite(raw) && raw >= 0
    ? Math.floor(raw)
    : null;
}

function workbookSheetCollection(
  app: WpsEtApplication,
  workbook: WpsEtWorkbook,
): WpsCountedCollection | null {
  return workbook.Worksheets
    ?? workbook.Sheets
    ?? app.Worksheets
    ?? app.Sheets
    ?? null;
}

function getCollectionItem(collection: WpsCountedCollection, key: string | number): DynamicValue {
  if (typeof collection.Item !== "function") {
    throw new Error("WPS worksheet collection does not expose Item().");
  }

  return collection.Item(key);
}

function listWorksheets(app: WpsEtApplication, workbook: WpsEtWorkbook): WpsEtWorksheet[] {
  const collection = workbookSheetCollection(app, workbook);
  if (!collection) {
    throw new Error("WPS workbook does not expose a worksheet collection.");
  }

  const count = getCount(collection);
  if (count === null) {
    throw new Error("WPS worksheet collection does not expose Count.");
  }

  const sheets: WpsEtWorksheet[] = [];
  for (let index = 1; index <= count; index += 1) {
    const sheet = asWorksheet(getCollectionItem(collection, index));
    if (sheet) sheets.push(sheet);
  }

  return sheets;
}

function getWorksheetByName(
  app: WpsEtApplication,
  workbook: WpsEtWorkbook,
  sheetName: string,
): WpsEtWorksheet {
  const collection = workbookSheetCollection(app, workbook);
  if (!collection) {
    throw new Error("WPS workbook does not expose a worksheet collection.");
  }

  const direct = asWorksheet(getCollectionItem(collection, sheetName));
  if (direct) return direct;

  throw new Error(`WPS worksheet not found: ${sheetName}`);
}

function getActiveWorksheet(app: WpsEtApplication, workbook: WpsEtWorkbook): WpsEtWorksheet {
  const active = asWorksheet(app.ActiveSheet);
  if (active) return active;

  const first = listWorksheets(app, workbook)[0];
  if (first) return first;

  throw new Error("WPS workbook has no worksheets.");
}

function requireWpsWorkbook(): { app: WpsEtApplication; workbook: WpsEtWorkbook } {
  const app = getWpsEtApplication();
  if (!app) {
    throw new Error("WPS ET Application is unavailable.");
  }

  const workbook = app.ActiveWorkbook;
  if (!workbook) {
    throw new Error("No active WPS workbook.");
  }

  return { app, workbook };
}

function normalizeAddress(address: string): string {
  return address.replace(/\$/gu, "");
}

function getRangeAddress(range: WpsEtRange, fallbackAddress: string): string {
  const address = asString(range.Address);
  return normalizeAddress(address ?? fallbackAddress);
}

function parseRectangularAddress(address: string): { rows: number; cols: number } | null {
  const clean = normalizeAddress(address.includes("!") ? address.split("!").at(-1) ?? address : address);
  if (!clean || /^\d+:\d+$/u.test(clean) || /^[A-Z]+:[A-Z]+$/iu.test(clean)) return null;

  const parts = clean.includes(":") ? clean.split(":") : [clean, clean];
  if (parts.length !== 2) return null;
  const startPart = parts[0];
  const endPart = parts[1];
  if (startPart === undefined || endPart === undefined) return null;

  try {
    const start = parseCell(startPart);
    const end = parseCell(endPart);
    return {
      rows: Math.abs(end.row - start.row) + 1,
      cols: Math.abs(end.col - start.col) + 1,
    };
  } catch {
    return null;
  }
}

function inferGridDims(raw: DynamicValue): { rows: number; cols: number } | null {
  if (isUnknownGrid(raw)) {
    return {
      rows: raw.length,
      cols: raw.reduce((max, row) => Math.max(max, row.length), 0),
    };
  }

  if (Array.isArray(raw)) {
    return { rows: 1, cols: raw.length };
  }

  if (raw !== undefined && raw !== null) {
    return { rows: 1, cols: 1 };
  }

  return null;
}

function rangeDimensions(range: WpsEtRange, address: string, rawValues: DynamicValue): { rows: number; cols: number } {
  const rowCount = getCount(range.Rows);
  const colCount = getCount(range.Columns);
  if (rowCount !== null && colCount !== null) {
    return { rows: rowCount, cols: colCount };
  }

  const fromAddress = parseRectangularAddress(address);
  if (fromAddress) return fromAddress;

  const fromGrid = inferGridDims(rawValues);
  if (fromGrid) return fromGrid;

  return { rows: 1, cols: 1 };
}

function normalizeGrid(raw: DynamicValue, rows: number, cols: number, emptyValue: DynamicValue): DynamicValue[][] {
  let grid: DynamicValue[][];

  if (isUnknownGrid(raw)) {
    grid = raw.map((row) => [...row]);
  } else if (isUnknownArray(raw)) {
    grid = rows === 1
      ? [[...raw]]
      : raw.map((value) => [value]);
  } else {
    grid = [[raw ?? emptyValue]];
  }

  const normalized: DynamicValue[][] = [];
  for (let row = 0; row < rows; row += 1) {
    const sourceRow = grid[row] ?? [];
    const normalizedRow: DynamicValue[] = [];
    for (let col = 0; col < cols; col += 1) {
      normalizedRow.push(sourceRow[col] ?? emptyValue);
    }
    normalized.push(normalizedRow);
  }

  return normalized;
}

function readRangeValue2(range: WpsEtRange): DynamicValue {
  if (range.Value2 !== undefined) return range.Value2;
  if (typeof range.Value === "function") return range.Value();
  return undefined;
}

function rangeSnapshot(sheet: WpsEtWorksheet, range: WpsEtRange, fallbackAddress: string): WpsRangeSnapshot {
  const rawValues = readRangeValue2(range);
  const address = getRangeAddress(range, fallbackAddress);
  const dimensions = rangeDimensions(range, address, rawValues);
  const values = normalizeGrid(rawValues, dimensions.rows, dimensions.cols, "");
  const metadataWarnings: string[] = [];

  let formulas: DynamicValue[][];
  if (range.Formula === undefined) {
    formulas = normalizeGrid(undefined, dimensions.rows, dimensions.cols, "");
    metadataWarnings.push("WPS Formula metadata was unavailable; formula listings may be incomplete.");
  } else {
    formulas = normalizeGrid(range.Formula, dimensions.rows, dimensions.cols, "");
  }

  let numberFormats: DynamicValue[][];
  if (range.NumberFormat === undefined) {
    numberFormats = normalizeGrid(undefined, dimensions.rows, dimensions.cols, "General");
    metadataWarnings.push("WPS NumberFormat metadata was unavailable; detailed format output may be incomplete.");
  } else {
    numberFormats = normalizeGrid(range.NumberFormat, dimensions.rows, dimensions.cols, "General");
  }

  return {
    sheetName: worksheetName(sheet),
    address,
    rows: dimensions.rows,
    cols: dimensions.cols,
    values,
    formulas,
    numberFormats,
    metadataWarnings,
  };
}

function getRangeForRef(
  app: WpsEtApplication,
  workbook: WpsEtWorkbook,
  ref: string,
): { sheet: WpsEtWorksheet; range: WpsEtRange; address: string } {
  const parsed = parseRangeRef(ref);
  const sheet = parsed.sheet
    ? getWorksheetByName(app, workbook, parsed.sheet)
    : getActiveWorksheet(app, workbook);

  if (typeof sheet.Range !== "function") {
    throw new Error(`WPS worksheet ${worksheetName(sheet)} does not expose Range().`);
  }

  const range = asRange(sheet.Range(parsed.address));
  if (!range) {
    throw new Error(`WPS Range("${parsed.address}") did not return a range object.`);
  }

  return { sheet, range, address: parsed.address };
}

function visibilityLabel(sheet: WpsEtWorksheet): string {
  const value = sheet.Visible ?? sheet.visible;
  if (value === true || value === -1 || value === "Visible" || value === "xlSheetVisible") return "Visible";
  if (value === false || value === 0 || value === "Hidden" || value === "xlSheetHidden") return "Hidden";
  if (value === 2 || value === "VeryHidden" || value === "xlSheetVeryHidden") return "VeryHidden";
  if (value === undefined || value === null) return "unknown visibility";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value);
  return "custom visibility";
}

function usedRangeSummary(sheet: WpsEtWorksheet): string {
  const used = sheet.UsedRange;
  if (!used) return "used range unavailable";

  const address = getRangeAddress(used, "A1");
  const rawValues = readRangeValue2(used);
  const dimensions = rangeDimensions(used, address, rawValues);
  return `${dimensions.rows} rows × ${dimensions.cols} cols (${address})`;
}

function selectionAddress(app: WpsEtApplication): string | null {
  const selection = asRange(app.Selection);
  if (!selection) return null;
  const address = asString(selection.Address);
  return address ? normalizeAddress(address) : null;
}

function explicitWpsOverviewOmissions(): string[] {
  return [
    "Headers, tables, named ranges, charts, PivotTables, shapes, and other workbook objects are not inventoried on WPS yet.",
  ];
}

function parseWorkbookOverviewParams(raw: DynamicValue): WorkbookOverviewParams {
  if (typeof raw !== "object" || raw === null) return {};
  const candidate = raw as { sheet?: DynamicValue };
  return typeof candidate.sheet === "string" && candidate.sheet.trim().length > 0
    ? { sheet: candidate.sheet }
    : {};
}

function parseReadRangeParams(raw: DynamicValue): ReadRangeParams {
  if (typeof raw !== "object" || raw === null) {
    throw new Error("read_range params must be an object.");
  }

  const candidate = raw as { range?: DynamicValue; mode?: DynamicValue };
  if (typeof candidate.range !== "string" || candidate.range.trim().length === 0) {
    throw new Error("read_range requires a non-empty range string.");
  }

  if (
    candidate.mode !== undefined &&
    candidate.mode !== "compact" &&
    candidate.mode !== "csv" &&
    candidate.mode !== "detailed"
  ) {
    throw new Error("read_range mode must be compact, csv, or detailed.");
  }

  return {
    range: candidate.range,
    ...(candidate.mode !== undefined ? { mode: candidate.mode } : {}),
  };
}

function parseWriteCellsParams(raw: DynamicValue): WriteCellsParams {
  if (typeof raw !== "object" || raw === null) {
    throw new Error("write_cells params must be an object.");
  }

  const candidate = raw as { start_cell?: DynamicValue; values?: DynamicValue; allow_overwrite?: DynamicValue };
  if (typeof candidate.start_cell !== "string" || candidate.start_cell.trim().length === 0) {
    throw new Error("write_cells requires a non-empty start_cell string.");
  }

  if (!isUnknownGrid(candidate.values)) {
    throw new Error("write_cells values must be a 2D array.");
  }

  return {
    start_cell: candidate.start_cell,
    values: candidate.values,
    allow_overwrite: candidate.allow_overwrite === true,
  };
}

export function executeWpsGetWorkbookOverview(
  _toolCallId: string,
  rawParams: DynamicValue,
): Promise<AgentToolResult<undefined>> {
  try {
    const params = parseWorkbookOverviewParams(rawParams);
    const { app, workbook } = requireWpsWorkbook();
    const lines = params.sheet
      ? buildWpsSheetDetail(app, workbook, params.sheet)
      : buildWpsOverview(app, workbook);

    return Promise.resolve({
      content: [{ type: "text", text: lines.join("\n") }],
      details: undefined,
    });
  } catch (error) {
    return Promise.resolve({
      content: [{ type: "text", text: `Error getting WPS workbook overview: ${getErrorMessage(error)}` }],
      details: undefined,
    });
  }
}

function buildWpsOverview(app: WpsEtApplication, workbook: WpsEtWorkbook): string[] {
  const sheets = listWorksheets(app, workbook);
  const activeName = app.ActiveSheet ? worksheetName(app.ActiveSheet) : null;
  const currentSelection = selectionAddress(app);

  const lines: string[] = [];
  lines.push(`## Workbook: ${workbookName(workbook)}`);
  lines.push("");
  lines.push(`### Sheets (${sheets.length})`);

  for (const [index, sheet] of sheets.entries()) {
    const visibility = visibilityLabel(sheet);
    const suffix = visibility === "Visible" ? "" : ` (${visibility})`;
    const activeMarker = activeName === worksheetName(sheet) ? " — active" : "";
    lines.push(`${index + 1}. **${worksheetName(sheet)}**${suffix} — ${usedRangeSummary(sheet)}${activeMarker}`);
  }

  if (activeName || currentSelection) {
    lines.push("");
    lines.push("### Current context");
    if (activeName) lines.push(`- Active sheet: **${activeName}**`);
    if (currentSelection) lines.push(`- Selection: ${currentSelection}`);
  }

  lines.push("");
  lines.push("### WPS Phase 2 limitations");
  for (const omission of explicitWpsOverviewOmissions()) {
    lines.push(`- ${omission}`);
  }

  return lines;
}

function buildWpsSheetDetail(
  app: WpsEtApplication,
  workbook: WpsEtWorkbook,
  sheetName: string,
): string[] {
  const sheet = getWorksheetByName(app, workbook, sheetName);
  const used = sheet.UsedRange;
  const lines: string[] = [];

  lines.push(`## Sheet: ${worksheetName(sheet)}${visibilityLabel(sheet) === "Visible" ? "" : ` (${visibilityLabel(sheet)})`}`);
  lines.push(`Dimensions: ${usedRangeSummary(sheet)}`);

  if (used) {
    const snapshot = rangeSnapshot(sheet, used, getRangeAddress(used, "A1"));
    const previewRows = snapshot.values.slice(0, 5).map((row) => row.slice(0, 5));
    if (previewRows.length > 0) {
      lines.push("");
      lines.push("### Data preview (top-left 5×5 max)");
      lines.push(formatAsMarkdownTable(previewRows));
    }
  }

  lines.push("");
  lines.push("### WPS Phase 2 limitations");
  for (const omission of explicitWpsOverviewOmissions()) {
    lines.push(`- ${omission}`);
  }

  return lines;
}

export function executeWpsReadRange(
  _toolCallId: string,
  rawParams: DynamicValue,
): Promise<AgentToolResult<ReadRangeCsvDetails | undefined>> {
  let rangeLabel = "(unknown range)";
  try {
    const params = parseReadRangeParams(rawParams);
    rangeLabel = params.range;
    const mode = params.mode ?? "compact";
    const { app, workbook } = requireWpsWorkbook();
    const target = getRangeForRef(app, workbook, params.range);
    const snapshot = rangeSnapshot(target.sheet, target.range, target.address);
    const fullAddress = qualifiedAddress(snapshot.sheetName, snapshot.address);
    const cellPart = snapshot.address.includes("!") ? snapshot.address.split("!").at(-1) ?? snapshot.address : snapshot.address;
    const startCell = cellPart.split(":")[0] ?? cellPart;

    if (mode === "csv") {
      return Promise.resolve(formatWpsCsvOutput(fullAddress, snapshot, startCell));
    }

    if (mode === "detailed") {
      return Promise.resolve(formatWpsDetailedOutput(fullAddress, snapshot, startCell));
    }

    return Promise.resolve(formatWpsCompactOutput(fullAddress, snapshot, startCell));
  } catch (error) {
    return Promise.resolve({
      content: [{ type: "text", text: `Error reading WPS range "${rangeLabel}": ${getErrorMessage(error)}` }],
      details: undefined,
    });
  }
}

function hasAnyNonEmptyCell(values: DynamicValue[][]): boolean {
  for (const row of values) {
    for (const value of row) {
      if (value !== null && value !== undefined && value !== "") return true;
    }
  }
  return false;
}

function formatAsExcelMarkdownTable(values: DynamicValue[][], startCell: string): string {
  if (!values || values.length === 0) return "(empty)";

  const start = parseCell(startCell);
  const numCols = Math.max(...values.map((row) => row.length));
  const header: DynamicValue[] = [""];
  for (let col = 0; col < numCols; col += 1) {
    header.push(colToLetter(start.col + col));
  }

  const rows: DynamicValue[][] = [header];
  for (let row = 0; row < values.length; row += 1) {
    const valueRow = values[row] ?? [];
    const renderedRow: DynamicValue[] = [start.row + row, ...valueRow];
    while (renderedRow.length < numCols + 1) renderedRow.push("");
    rows.push(renderedRow);
  }

  return formatAsMarkdownTable(rows);
}

function appendMetadataWarnings(lines: string[], warnings: readonly string[]): void {
  if (warnings.length === 0) return;
  lines.push("");
  lines.push("### WPS metadata notes");
  for (const warning of warnings) {
    lines.push(`- ${warning}`);
  }
}

function formatWpsCompactOutput(
  address: string,
  snapshot: WpsRangeSnapshot,
  startCell: string,
): AgentToolResult<undefined> {
  const lines: string[] = [];
  const formulas = extractFormulas(snapshot.formulas, startCell);
  const errors = findErrors(snapshot.values, startCell);
  const hasValues = hasAnyNonEmptyCell(snapshot.values);

  lines.push(`**${address}** (${snapshot.rows}×${snapshot.cols})`);

  if (!hasValues && formulas.length === 0 && errors.length === 0) {
    lines.push("");
    lines.push("_All cells are empty._");
    appendMetadataWarnings(lines, snapshot.metadataWarnings);
    return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
  }

  lines.push("");
  lines.push(formatAsExcelMarkdownTable(snapshot.values, startCell));

  if (formulas.length > 0) {
    lines.push("");
    lines.push(`**Formulas:** ${formulas.join(", ")}`);
  }

  if (errors.length > 0) {
    lines.push("");
    lines.push(`⚠️ **Errors:** ${errors.map((error) => `${error.address}=${error.error}`).join(", ")}`);
  }

  appendMetadataWarnings(lines, snapshot.metadataWarnings);
  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}

function formatWpsDetailedOutput(
  address: string,
  snapshot: WpsRangeSnapshot,
  startCell: string,
): AgentToolResult<undefined> {
  const lines: string[] = [];
  const formulas = extractFormulas(snapshot.formulas, startCell);
  const errors = findErrors(snapshot.values, startCell);
  const hasValues = hasAnyNonEmptyCell(snapshot.values);

  lines.push(`**${address}** (${snapshot.rows}×${snapshot.cols})`);

  if (!hasValues && formulas.length === 0 && errors.length === 0) {
    lines.push("");
    lines.push("_All cells are empty._");
    appendMetadataWarnings(lines, snapshot.metadataWarnings);
    return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
  }

  lines.push("");
  lines.push("### Values");
  lines.push(formatAsExcelMarkdownTable(snapshot.values, startCell));

  if (formulas.length > 0) {
    lines.push("");
    lines.push("### Formulas");
    for (const formula of formulas) {
      lines.push(`- ${formula}`);
    }
  }

  const formatMap = new Map<string, string[]>();
  const start = parseCell(startCell);
  for (let row = 0; row < snapshot.numberFormats.length; row += 1) {
    const numberFormatRow = snapshot.numberFormats[row] ?? [];
    for (let col = 0; col < numberFormatRow.length; col += 1) {
      const format = numberFormatRow[col];
      if (typeof format === "string" && format !== "" && format !== "General") {
        const addressForCell = `${colToLetter(start.col + col)}${start.row + row}`;
        const existing = formatMap.get(format) ?? [];
        existing.push(addressForCell);
        formatMap.set(format, existing);
      }
    }
  }

  if (formatMap.size > 0) {
    lines.push("");
    lines.push("### Number Formats");
    for (const [format, cells] of formatMap) {
      lines.push(`- \`${format}\` → ${cells.join(", ")}`);
    }
  }

  if (errors.length > 0) {
    lines.push("");
    lines.push("### ⚠️ Errors");
    for (const error of errors) {
      lines.push(`- ${error.address}: ${error.error}`);
    }
  }

  appendMetadataWarnings(lines, snapshot.metadataWarnings);
  return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
}

function toCsvField(value: DynamicValue): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") {
    return value.includes(",") || value.includes('"') || value.includes("\n") || value.includes("\r")
      ? `"${value.replace(/"/gu, '""')}"`
      : value;
  }
  const raw = typeof value === "number" || typeof value === "boolean" ? String(value) : JSON.stringify(value);
  const serialized = raw ?? "";
  return /[",\n\r]/u.test(serialized) ? `"${serialized.replace(/"/gu, '""')}"` : serialized;
}

function valuesToCsv(values: DynamicValue[][]): string {
  if (!values || values.length === 0) return "";
  return values.map((row) => row.map((value) => toCsvField(value)).join(",")).join("\n");
}

function formatWpsCsvOutput(
  address: string,
  snapshot: WpsRangeSnapshot,
  startCell: string,
): AgentToolResult<ReadRangeCsvDetails | undefined> {
  const lines: string[] = [];
  lines.push(`**${address}** (${snapshot.rows}×${snapshot.cols})`);
  lines.push("");

  const csv = valuesToCsv(snapshot.values);
  if (!csv) {
    lines.push("(empty)");
    appendMetadataWarnings(lines, snapshot.metadataWarnings);
    return { content: [{ type: "text", text: lines.join("\n") }], details: undefined };
  }

  lines.push("```csv");
  lines.push(csv);
  lines.push("```");
  appendMetadataWarnings(lines, snapshot.metadataWarnings);

  const start = parseCell(startCell);
  return {
    content: [{ type: "text", text: lines.join("\n") }],
    details: {
      kind: "read_range_csv",
      startCol: start.col,
      startRow: start.row,
      values: snapshot.values,
      csv,
    },
  };
}

function findInvalidFormulas(values: DynamicValue[][], startCell: string): InvalidFormula[] {
  const start = parseCell(startCell);
  const invalid: InvalidFormula[] = [];

  for (let row = 0; row < values.length; row += 1) {
    const valueRow = values[row] ?? [];
    for (let col = 0; col < valueRow.length; col += 1) {
      const value = valueRow[col];
      if (typeof value === "string" && value.startsWith("=")) {
        const reason = validateFormula(value);
        if (reason) {
          invalid.push({
            address: `${colToLetter(start.col + col)}${start.row + row}`,
            formula: value,
            reason,
          });
        }
      }
    }
  }

  return invalid;
}

function writeWpsRange(range: WpsEtRange, values: DynamicValue[][]): void {
  const containsFormula = values.some((row) => row.some((value) => typeof value === "string" && value.startsWith("=")));
  if (containsFormula) {
    range.Formula = values;
    return;
  }

  if (typeof range.Value === "function" && range.Value2 === undefined) {
    range.Value(undefined, values);
    return;
  }

  range.Value2 = values;
}

export function executeWpsWriteCells(
  _toolCallId: string,
  rawParams: DynamicValue,
): Promise<AgentToolResult<WriteCellsDetails>> {
  try {
    const params = parseWriteCellsParams(rawParams);
    if (params.values.length === 0) {
      return Promise.resolve({
        content: [{ type: "text", text: "Error: values array is empty." }],
        details: { kind: "write_cells", blocked: false },
      });
    }

    const { padded, rows, cols } = padValues(params.values);
    const startCellRef = params.start_cell.includes("!")
      ? params.start_cell.split("!").at(-1) ?? params.start_cell
      : params.start_cell;

    if (startCellRef.includes(":")) {
      return Promise.resolve({
        content: [{ type: "text", text: "Error: start_cell must be a single cell (e.g. \"A1\")." }],
        details: { kind: "write_cells", blocked: false },
      });
    }

    let invalidFormulas: InvalidFormula[] = [];
    try {
      invalidFormulas = findInvalidFormulas(padded, startCellRef);
    } catch {
      return Promise.resolve({
        content: [{ type: "text", text: `Error: invalid start_cell "${params.start_cell}".` }],
        details: { kind: "write_cells", blocked: false },
      });
    }

    if (invalidFormulas.length > 0) {
      const lines: string[] = [];
      lines.push("⛔ **Write blocked** — invalid formula syntax detected:");
      for (const invalid of invalidFormulas) {
        lines.push(`- ${invalid.address}: ${invalid.formula} (${invalid.reason})`);
      }
      lines.push("");
      lines.push("Fix the formulas and retry.");
      return Promise.resolve({
        content: [{ type: "text", text: lines.join("\n") }],
        details: { kind: "write_cells", blocked: true },
      });
    }

    const result = writeWpsCells(params, padded, rows, cols, startCellRef);
    return Promise.resolve(result.blocked ? formatWpsBlockedWrite(result) : formatWpsSuccessWrite(result, rows, cols));
  } catch (error) {
    return Promise.resolve({
      content: [{ type: "text", text: `Error writing WPS cells: ${getErrorMessage(error)}` }],
      details: { kind: "write_cells", blocked: false },
    });
  }
}

function writeWpsCells(
  params: WriteCellsParams,
  padded: DynamicValue[][],
  rows: number,
  cols: number,
  startCellRef: string,
): WriteCellsResult {
  const { app, workbook } = requireWpsWorkbook();
  const { sheet } = getRangeForRef(app, workbook, params.start_cell);
  const rangeAddress = computeRangeAddress(startCellRef, rows, cols);
  if (typeof sheet.Range !== "function") {
    throw new Error(`WPS worksheet ${worksheetName(sheet)} does not expose Range().`);
  }

  const targetRange = asRange(sheet.Range(rangeAddress));
  if (!targetRange) {
    throw new Error(`WPS Range("${rangeAddress}") did not return a range object.`);
  }

  const before = rangeSnapshot(sheet, targetRange, rangeAddress);
  if (!params.allow_overwrite) {
    const occupiedCount = countOccupiedCells(before.values, before.formulas);
    if (occupiedCount > 0) {
      return {
        blocked: true,
        sheetName: before.sheetName,
        address: before.address,
        existingCount: occupiedCount,
        existingValues: before.values,
      };
    }
  }

  writeWpsRange(targetRange, padded);
  const readBack = rangeSnapshot(sheet, targetRange, rangeAddress);

  return {
    blocked: false,
    sheetName: readBack.sheetName,
    address: readBack.address,
    beforeValues: before.values,
    beforeFormulas: before.formulas,
    readBackValues: readBack.values,
    readBackFormulas: readBack.formulas,
  };
}

function formatWpsBlockedWrite(result: BlockedWriteCellsResult): AgentToolResult<WriteCellsDetails> {
  const fullAddress = qualifiedAddress(result.sheetName, result.address);
  const lines: string[] = [];

  lines.push(`⛔ **Write blocked** — ${fullAddress} contains ${result.existingCount} non-empty cell(s).`);
  lines.push("");
  lines.push("**Existing data:**");
  lines.push(result.existingCount > 0 ? formatAsMarkdownTable(result.existingValues) : "(empty)");
  lines.push("");
  lines.push("To overwrite, confirm with the user and retry with `allow_overwrite: true`.");
  lines.push("");
  lines.push(WPS_NO_BACKUP_NOTICE);

  return {
    content: [{ type: "text", text: lines.join("\n") }],
    details: {
      kind: "write_cells",
      blocked: true,
      address: fullAddress,
      existingCount: result.existingCount,
      recovery: { status: "not_available", reason: WPS_NO_BACKUP_REASON },
    },
  };
}

const VERIFIED_VALUES_PREVIEW_ROWS = 8;
const VERIFIED_VALUES_PREVIEW_COLS = 6;

function buildVerifiedValuesPreview(values: DynamicValue[][]): {
  values: DynamicValue[][];
  totalRows: number;
  totalCols: number;
  shownRows: number;
  shownCols: number;
  omittedRows: number;
  omittedCols: number;
  truncated: boolean;
} {
  const totalRows = values.length;
  const totalCols = values.reduce((max, row) => Math.max(max, row.length), 0);
  const shownRows = Math.min(totalRows, VERIFIED_VALUES_PREVIEW_ROWS);
  const shownCols = Math.min(totalCols, VERIFIED_VALUES_PREVIEW_COLS);
  const previewValues = values.slice(0, shownRows).map((row) => row.slice(0, shownCols));
  const omittedRows = Math.max(totalRows - shownRows, 0);
  const omittedCols = Math.max(totalCols - shownCols, 0);

  return {
    values: previewValues,
    totalRows,
    totalCols,
    shownRows,
    shownCols,
    omittedRows,
    omittedCols,
    truncated: omittedRows > 0 || omittedCols > 0,
  };
}

function formatWpsSuccessWrite(
  result: SuccessWriteCellsResult,
  rows: number,
  cols: number,
): AgentToolResult<WriteCellsDetails> {
  const fullAddress = qualifiedAddress(result.sheetName, result.address);
  const cellPart = result.address.includes("!") ? result.address.split("!").at(-1) ?? result.address : result.address;
  const startCell = cellPart.split(":")[0] ?? cellPart;
  const lines: string[] = [];
  lines.push(`Written to **${fullAddress}** (${rows}×${cols})`);

  const errors = findErrors(result.readBackValues, startCell);
  if (errors.length > 0) {
    const start = parseCell(startCell);
    for (const error of errors) {
      const errorCell = parseCell(error.address);
      const row = errorCell.row - start.row;
      const col = errorCell.col - start.col;
      const formulaRow = result.readBackFormulas[row];
      if (row >= 0 && col >= 0 && formulaRow !== undefined && col < formulaRow.length) {
        const formula = formulaRow[col];
        if (typeof formula === "string") error.formula = formula;
      }
    }

    lines.push("");
    lines.push(`⚠️ **${errors.length} formula error(s):**`);
    for (const error of errors) {
      lines.push(`- ${error.address}: ${error.error}${error.formula ? ` (formula: ${error.formula})` : ""}`);
    }
    lines.push("");
    lines.push("Review and fix with another write_cells call.");
  } else {
    const preview = buildVerifiedValuesPreview(result.readBackValues);
    lines.push("");
    if (preview.truncated) {
      lines.push(`**Verified values (preview ${preview.shownRows}×${preview.shownCols} of ${preview.totalRows}×${preview.totalCols}):**`);
    } else {
      lines.push("**Verified values:**");
    }
    lines.push(formatAsMarkdownTable(preview.values));

    if (preview.truncated) {
      const omissions: string[] = [];
      if (preview.omittedRows > 0) {
        omissions.push(`${preview.omittedRows} more row${preview.omittedRows === 1 ? "" : "s"}`);
      }
      if (preview.omittedCols > 0) {
        omissions.push(`${preview.omittedCols} more column${preview.omittedCols === 1 ? "" : "s"}`);
      }
      lines.push("");
      lines.push(`_Showing preview only (${omissions.join(" and ")})._`);
      lines.push("_Use `read_range` for full verification if needed._");
    }
  }

  const changes = buildWorkbookCellChangeSummary({
    sheetName: result.sheetName,
    startCell,
    beforeValues: result.beforeValues,
    beforeFormulas: result.beforeFormulas,
    afterValues: result.readBackValues,
    afterFormulas: result.readBackFormulas,
  });

  if (changes.changedCount > 0) {
    lines.push("");
    lines.push(`Changed cell(s): ${changes.changedCount}.`);
  }

  lines.push("");
  lines.push(WPS_NO_BACKUP_NOTICE);

  return {
    content: [{ type: "text", text: lines.join("\n") }],
    details: {
      kind: "write_cells",
      blocked: false,
      address: fullAddress,
      formulaErrorCount: errors.length,
      changes,
      recovery: { status: "not_available", reason: WPS_NO_BACKUP_REASON },
    },
  };
}
