import { isExcelError } from "../../utils/format.js";
import {
  cellAddress,
  colToLetter,
  parseCell,
  qualifiedAddress,
} from "../../excel/helpers.js";
import { buildWorkbookCellChangeSummary, type WorkbookCellChangeSummary } from "../../audit/cell-diff.js";
import { MAX_RECOVERY_CELLS } from "../../workbook/recovery/constants.js";
import { CHECKPOINT_SKIPPED_REASON, tooLargeToBackUpReason } from "../recovery-metadata.js";

export const FULL_VERIFY_CELL_LIMIT = MAX_RECOVERY_CELLS;

const SAMPLE_CELLS_PER_END = 1_000;

const PER_CELL_PROBE_LIMIT = 500;

const LISTED_CELL_LIMIT = 8;

export type CellScalar = string | number | boolean | null | undefined;

export function toCellGrid(grid: readonly (readonly unknown[])[]): CellScalar[][] {
  return grid.map((row) => row.map((value) => {
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
    if (value === null || value === undefined) return null;
    return JSON.stringify(value) ?? null;
  }));
}


export interface WriteBlock {
  startRow: number;
  startCol: number;
  rowCount: number;
  columnCount: number;
}

export function writeBlockFromAddress(address: string): WriteBlock {
  const cellPart = address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address;
  const [startText = cellPart, endText = startText] = cellPart.split(":");
  const start = parseCell(startText);
  const end = parseCell(endText);
  return {
    startRow: start.row,
    startCol: start.col,
    rowCount: end.row - start.row + 1,
    columnCount: end.col - start.col + 1,
  };
}

export function blockAddress(block: WriteBlock, rowOffset = 0, rowCount = block.rowCount): string {
  const top = block.startRow + rowOffset;
  const left = colToLetter(block.startCol);
  const right = colToLetter(block.startCol + block.columnCount - 1);
  const bottom = top + rowCount - 1;
  return top === bottom && block.columnCount === 1 ? `${left}${top}` : `${left}${top}:${right}${bottom}`;
}

function cellAt(block: WriteBlock, row: number, col: number): string {
  return cellAddress(block.startCol + col, block.startRow + row);
}

function cellCount(block: WriteBlock): number {
  return block.rowCount * block.columnCount;
}

/**
 * What the tool sends to one cell. `skip` means the cell is left as Excel has
 * it (Office.js treats `null` that way); `formula` is any string starting "=".
 */
export type SentCell =
  | { kind: "skip" }
  | { kind: "empty" }
  | { kind: "formula"; formula: string }
  | { kind: "text"; text: string; literal: boolean }
  | { kind: "number"; value: number }
  | { kind: "boolean"; value: boolean };

export function classifySentValue(value: CellScalar): SentCell {
  if (value === null || value === undefined) return { kind: "skip" };
  if (typeof value === "number") return { kind: "number", value };
  if (typeof value === "boolean") return { kind: "boolean", value };
  if (typeof value !== "string") return { kind: "skip" };
  if (value === "") return { kind: "empty" };
  if (value.startsWith("=")) return { kind: "formula", formula: value };
  if (value.startsWith("'")) return { kind: "text", text: value.slice(1), literal: true };
  return { kind: "text", text: value, literal: false };
}

export type SentGrid = (row: number, col: number) => SentCell;

function sendsContent(cell: SentCell): boolean {
  return cell.kind !== "skip" && cell.kind !== "empty";
}

export interface MergeConflict {
  mergedArea: string;
  cells: string[];
}

export interface SpillConflict {
  cell: string;
  anchor: string;
  anchorFormula: string;
}

export interface WriteTargetInspection {
  sheetName: string;
  calculationMode: string | null;
  sheetProtected: boolean;
  targetUnlocked: boolean;
  before: { values: CellScalar[][]; formulas: CellScalar[][] } | null;
  occupiedAddress: string | null;
  occupiedCount: number;
  mergeConflicts: MergeConflict[];
  spillConflicts: SpillConflict[];
  tableNames: string[] | null;
  checkNotes: string[];
}

export function excelApiSupported(version: string): boolean {
  if (typeof Office === "undefined") return true;
  return Office.context?.requirements?.isSetSupported?.("ExcelApi", version) ?? true;
}

function stripSheet(address: string): string {
  return address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address;
}

export function countOccupiedCells(values: unknown[][], formulas: unknown[][]): number {
  let count = 0;
  for (let r = 0; r < values.length; r++) {
    const valueRow = values[r] ?? [];
    for (let c = 0; c < valueRow.length; c++) {
      const value = valueRow[c];
      const formula = formulas[r]?.[c];
      const hasValue = value !== null && value !== undefined && value !== "";
      const hasFormula = typeof formula === "string" && formula.startsWith("=");
      if (hasValue || hasFormula) count += 1;
    }
  }
  return count;
}

export async function inspectWriteTarget(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  block: WriteBlock,
  sent: SentGrid,
): Promise<WriteTargetInspection> {
  const target = sheet.getRange(blockAddress(block));
  const small = cellCount(block) <= FULL_VERIFY_CELL_LIMIT;
  const checkNotes: string[] = [];

  sheet.load("name");
  sheet.protection.load("protected");
  target.format.protection.load("locked");
  const application = context.workbook.application;
  application.load("calculationMode");

  if (small) target.load("values,formulas");
  const used = small ? null : target.getUsedRangeOrNullObject(true);
  used?.load("address,cellCount");

  const canReadMerges = excelApiSupported("1.13");
  const merges = canReadMerges
    ? sheet.getRange(`A1:${colToLetter(block.startCol + block.columnCount - 1)}${block.startRow + block.rowCount - 1}`)
      .getMergedAreasOrNullObject()
    : null;
  merges?.load("areas/items/address,areas/items/rowCount,areas/items/columnCount");
  if (!canReadMerges) checkNotes.push("merged cells were not checked (needs ExcelApi 1.13)");

  const tables = excelApiSupported("1.9") ? target.getTables(false) : null;
  tables?.load("items/name");

  const canReadSpills = excelApiSupported("1.12");
  if (canReadSpills) target.load("hasSpill");
  else checkNotes.push("spilled ranges were not checked (needs ExcelApi 1.12)");

  await context.sync();

  const mergeConflicts: MergeConflict[] = [];
  if (merges && !merges.isNullObject) {
    for (const area of merges.areas.items) {
      const topLeftText = stripSheet(area.address).split(":")[0] ?? "";
      const topLeft = parseCell(topLeftText);
      const cells: string[] = [];
      for (let row = topLeft.row; row < topLeft.row + area.rowCount; row += 1) {
        for (let col = topLeft.col; col < topLeft.col + area.columnCount; col += 1) {
          if (row === topLeft.row && col === topLeft.col) continue;
          const r = row - block.startRow;
          const c = col - block.startCol;
          if (r < 0 || c < 0 || r >= block.rowCount || c >= block.columnCount) continue;
          if (sendsContent(sent(r, c))) cells.push(cellAddress(col, row));
        }
      }
      if (cells.length > 0) {
        const bottomRight = cellAddress(topLeft.col + area.columnCount - 1, topLeft.row + area.rowCount - 1);
        mergeConflicts.push({ mergedArea: `${topLeftText}:${bottomRight}`, cells });
      }
    }
  }

  const spillConflicts: SpillConflict[] = [];
  if (canReadSpills && target.hasSpill !== false) {
    const probes: Array<{ cell: string; parent: Excel.Range }> = [];
    let candidates = 0;
    for (let r = 0; r < block.rowCount; r += 1) {
      for (let c = 0; c < block.columnCount; c += 1) {
        if (!sendsContent(sent(r, c))) continue;
        candidates += 1;
        if (probes.length >= PER_CELL_PROBE_LIMIT) continue;
        const cell = cellAt(block, r, c);
        const parent = sheet.getRange(cell).getSpillParentOrNullObject();
        parent.load("address,formulas");
        probes.push({ cell, parent });
      }
    }
    if (candidates > probes.length) {
      checkNotes.push(`spilled ranges were checked for the first ${probes.length} of ${candidates} cells only`);
    }
    if (probes.length > 0) await context.sync();
    for (const probe of probes) {
      if (probe.parent.isNullObject) continue;
      const anchor = stripSheet(probe.parent.address);
      if (anchor === probe.cell) continue;
      const formula = toCellGrid(probe.parent.formulas)[0]?.[0];
      spillConflicts.push({ cell: probe.cell, anchor, anchorFormula: typeof formula === "string" ? formula : "" });
    }
  }

  const before = small ? { values: toCellGrid(target.values), formulas: toCellGrid(target.formulas) } : null;
  const occupiedCount = before
    ? countOccupiedCells(before.values, before.formulas)
    : used && !used.isNullObject ? used.cellCount : 0;
  const occupiedAddress = before
    ? null
    : used && !used.isNullObject ? stripSheet(used.address) : null;

  return {
    sheetName: sheet.name,
    calculationMode: application.calculationMode ?? null,
    sheetProtected: sheet.protection.protected,
    targetUnlocked: target.format.protection.locked === false,
    before,
    occupiedAddress,
    occupiedCount,
    mergeConflicts,
    spillConflicts,
    tableNames: tables ? tables.items.map((table) => table.name) : null,
    checkNotes,
  };
}

export function listCells(cells: readonly string[], limit = LISTED_CELL_LIMIT): string {
  const shown = cells.slice(0, limit).join(", ");
  return cells.length > limit ? `${shown} and ${cells.length - limit} more` : shown;
}

export function describeWriteObstacle(
  inspection: WriteTargetInspection,
  qualify: (address: string) => string,
): string | null {
  if (inspection.sheetProtected && !inspection.targetUnlocked) {
    return `Sheet "${inspection.sheetName}" is protected and the target cells are locked, so Excel would refuse the write. ` +
      "Pi cannot unprotect it: ask the user to unprotect the sheet (Review → Unprotect Sheet) or unlock these cells.";
  }
  if (inspection.mergeConflicts.length > 0) {
    const lines = inspection.mergeConflicts.map((conflict) =>
      `- ${qualify(conflict.mergedArea)} is merged; only its top-left cell holds a value, so ${listCells(conflict.cells)} would be dropped.`);
    return [
      "The target overlaps merged cells. Excel keeps only the top-left value of a merged area and silently drops the rest:",
      ...lines,
      "Write each value to the merged area's top-left cell, or unmerge first (format_cells with merge: \"unmerge\").",
    ].join("\n");
  }
  if (inspection.spillConflicts.length > 0) {
    const byAnchor = new Map<string, SpillConflict[]>();
    for (const conflict of inspection.spillConflicts) {
      byAnchor.set(conflict.anchor, [...(byAnchor.get(conflict.anchor) ?? []), conflict]);
    }
    const lines = [...byAnchor.entries()].map(([anchor, conflicts]) =>
      `- ${listCells(conflicts.map((conflict) => conflict.cell))} ${conflicts.length === 1 ? "is" : "are"} spilled from ${qualify(anchor)}` +
      `${conflicts[0]?.anchorFormula ? ` (${conflicts[0].anchorFormula})` : ""}; writing there turns ${anchor} into #SPILL!.`);
    return [
      "The target is inside a spilled range (output of a dynamic array formula):",
      ...lines,
      "Change the anchor formula instead, or write outside the spilled range.",
    ].join("\n");
  }
  return null;
}

export interface CellIssue {
  address: string;
  sent: string;
  got: string;
  shown?: string;
}

export interface FormulaErrorCell {
  address: string;
  error: string;
  formula?: string;
}

export interface WriteVerification {
  mode: "full" | "sampled" | "failed";
  values: CellScalar[][] | null;
  formulas: CellScalar[][] | null;
  notWritten: CellIssue[];
  notWrittenCount: number;
  converted: CellIssue[];
  altered: CellIssue[];
  formulaErrors: FormulaErrorCell[];
  formulaErrorCount: number;
  invalidCells: string[];
  invalidSummary: string | null;
  invalidRule: string | null;
  tableExtensions: Array<{ name: string; address: string; newHeaders: string[] }>;
  manualCalculation: string | null;
  calculatedWritten: boolean;
  scopeNote: string | null;
  sampleRows: Array<{ rowOffset: number; values: CellScalar[][]; formulas: CellScalar[][] }>;
}

export type WriteReadBack =
  | { stopped: true; verification: null; verifyError: null }
  | { stopped: false; verification: WriteVerification; verifyError: null }
  | { stopped: false; verification: null; verifyError: string };

function preview(value: CellScalar): string {
  if (value === null || value === undefined || value === "") return "(empty)";
  if (typeof value === "string") return JSON.stringify(value.length > 40 ? `${value.slice(0, 39)}…` : value);
  return String(value);
}

function sentPreview(cell: SentCell): string {
  switch (cell.kind) {
    case "formula": return cell.formula;
    case "text": return JSON.stringify(cell.literal ? `'${cell.text}` : cell.text);
    case "number":
    case "boolean": return String(cell.value);
    default: return "(empty)";
  }
}

function isEmptyCell(value: CellScalar, formula: CellScalar): boolean {
  const hasFormula = typeof formula === "string" && formula.startsWith("=");
  return !hasFormula && (value === null || value === undefined || value === "");
}

function numbersMatch(a: number, b: number): boolean {
  if (a === b) return true;
  const scale = Math.max(Math.abs(a), Math.abs(b));
  return Math.abs(a - b) <= scale * 1e-14;
}

function roundTripsAsNumber(text: string, value: number): boolean {
  const trimmed = text.trim();
  return trimmed !== "" && String(value) === trimmed;
}

type CellOutcome = "ok" | "not_written" | "converted" | "altered";

function compareCell(sent: SentCell, value: CellScalar, formula: CellScalar): CellOutcome {
  switch (sent.kind) {
    case "skip":
      return "ok";
    case "empty":
      return isEmptyCell(value, formula) ? "ok" : "not_written";
    case "formula":
      return typeof formula === "string" && formula.startsWith("=") ? "ok" : "not_written";
    case "text": {
      if (isEmptyCell(value, formula)) return "not_written";
      if (typeof value === "string") return value === sent.text ? "ok" : "altered";
      if (typeof value === "number" && !sent.literal && roundTripsAsNumber(sent.text, value)) return "ok";
      if (typeof value === "boolean" && !sent.literal && sent.text.trim().toUpperCase() === String(value).toUpperCase()) return "ok";
      return "converted";
    }
    case "number":
      if (isEmptyCell(value, formula)) return "not_written";
      return typeof value === "number" && numbersMatch(value, sent.value) ? "ok" : "altered";
    case "boolean":
      if (isEmptyCell(value, formula)) return "not_written";
      return value === sent.value ? "ok" : "altered";
  }
}

interface DiffAccumulator {
  notWritten: CellIssue[];
  converted: CellIssue[];
  altered: CellIssue[];
  formulaErrors: FormulaErrorCell[];
}

function diffRows(
  block: WriteBlock,
  sent: SentGrid,
  rowOffset: number,
  values: CellScalar[][],
  formulas: CellScalar[][],
  into: DiffAccumulator,
): void {
  for (let r = 0; r < values.length; r += 1) {
    const valueRow = values[r] ?? [];
    for (let c = 0; c < block.columnCount; c += 1) {
      const row = rowOffset + r;
      const value = valueRow[c];
      const formula = formulas[r]?.[c];
      const sentCell = sent(row, c);
      const address = cellAt(block, row, c);
      if (isExcelError(value)) {
        into.formulaErrors.push({
          address,
          error: value,
          ...(typeof formula === "string" && formula.startsWith("=") ? { formula } : {}),
        });
      }
      const outcome = compareCell(sentCell, value, formula);
      if (outcome === "ok") continue;
      const issue: CellIssue = { address, sent: sentPreview(sentCell), got: preview(value) };
      if (outcome === "not_written") into.notWritten.push(issue);
      else if (outcome === "converted") into.converted.push(issue);
      else into.altered.push(issue);
    }
  }
}

export function diffWrittenGrid(
  block: WriteBlock,
  sent: SentGrid,
  values: CellScalar[][],
  formulas: CellScalar[][],
): WriteVerification {
  const acc: DiffAccumulator = { notWritten: [], converted: [], altered: [], formulaErrors: [] };
  diffRows(block, sent, 0, values, formulas, acc);
  return {
    mode: "full",
    values,
    formulas,
    notWritten: acc.notWritten,
    notWrittenCount: acc.notWritten.length,
    converted: acc.converted,
    altered: acc.altered,
    formulaErrors: acc.formulaErrors,
    formulaErrorCount: acc.formulaErrors.length,
    invalidCells: [],
    invalidSummary: null,
    tableExtensions: [],
    invalidRule: null,
    manualCalculation: null,
    calculatedWritten: false,
    scopeNote: null,
    sampleRows: [],
  };
}

export interface VerifyOptions {
  calculationMode: string | null;
  sentFormulas: boolean;
  sentEmptyCount: number;
  tablesBefore: string[] | null;
  sentChars: number;
}

const FULL_VERIFY_CHAR_LIMIT = 1_500_000;

export async function verifyWrittenRange(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  block: WriteBlock,
  sent: SentGrid,
  options: VerifyOptions,
): Promise<WriteVerification> {
  const target = sheet.getRange(blockAddress(block));
  const manual = options.calculationMode !== null && options.calculationMode !== "Automatic"
    ? options.calculationMode
    : null;
  const calculatedWritten = manual !== null && options.sentFormulas;
  if (calculatedWritten) {
    target.calculate();
    await context.sync();
  }

  const acc: DiffAccumulator = { notWritten: [], converted: [], altered: [], formulaErrors: [] };
  const verification: WriteVerification = {
    mode: "full",
    values: null,
    formulas: null,
    notWritten: acc.notWritten,
    notWrittenCount: 0,
    converted: acc.converted,
    altered: acc.altered,
    formulaErrors: acc.formulaErrors,
    formulaErrorCount: 0,
    invalidCells: [],
    invalidSummary: null,
    invalidRule: null,
    tableExtensions: [],
    manualCalculation: manual,
    calculatedWritten,
    scopeNote: null,
    sampleRows: [],
  };

  const canReadValidation = excelApiSupported("1.8");
  if (canReadValidation) target.dataValidation.load("valid");
  const tablesAfter = options.tablesBefore ? target.getTables(false) : null;
  tablesAfter?.load("items/name");

  if (cellCount(block) <= FULL_VERIFY_CELL_LIMIT && options.sentChars <= FULL_VERIFY_CHAR_LIMIT) {
    target.load("values,formulas");
    await context.sync();
    const values = toCellGrid(target.values);
    const formulas = toCellGrid(target.formulas);
    verification.values = values;
    verification.formulas = formulas;
    diffRows(block, sent, 0, values, formulas, acc);
    verification.notWrittenCount = acc.notWritten.length;
    verification.formulaErrorCount = acc.formulaErrors.length;
  } else {
    verification.mode = "sampled";
    const sampleRowCount = Math.max(1, Math.min(block.rowCount, Math.floor(SAMPLE_CELLS_PER_END / block.columnCount)));
    const head = sheet.getRange(blockAddress(block, 0, sampleRowCount));
    head.load("values,formulas");
    const tailOffset = Math.max(sampleRowCount, block.rowCount - sampleRowCount);
    const tailRows = block.rowCount - tailOffset;
    const tail = tailRows > 0 ? sheet.getRange(blockAddress(block, tailOffset, tailRows)) : null;
    tail?.load("values,formulas");
    const errorCells = target.getSpecialCellsOrNullObject("Formulas", "Errors");
    errorCells.load("address,cellCount");
    const blankCells = target.getSpecialCellsOrNullObject("Blanks");
    blankCells.load("address,cellCount");
    await context.sync();

    const headRows = { rowOffset: 0, values: toCellGrid(head.values), formulas: toCellGrid(head.formulas) };
    diffRows(block, sent, 0, headRows.values, headRows.formulas, acc);
    verification.sampleRows.push(headRows);
    if (tail) {
      const tailSample = { rowOffset: tailOffset, values: toCellGrid(tail.values), formulas: toCellGrid(tail.formulas) };
      diffRows(block, sent, tailOffset, tailSample.values, tailSample.formulas, acc);
      verification.sampleRows.push(tailSample);
    }

    const sampledErrors = acc.formulaErrors.length;
    const scannedErrors = errorCells.isNullObject ? 0 : errorCells.cellCount;
    verification.formulaErrorCount = Math.max(sampledErrors, scannedErrors);
    const blankCount = blankCells.isNullObject ? 0 : blankCells.cellCount;
    const missingBlanks = Math.max(0, blankCount - options.sentEmptyCount);
    verification.notWrittenCount = Math.max(acc.notWritten.length, missingBlanks);

    const scanned = [`read back rows ${block.startRow}–${block.startRow + sampleRowCount - 1}`];
    if (tail) scanned[0] += ` and ${block.startRow + tailOffset}–${block.startRow + block.rowCount - 1}`;
    scanned.push(`scanned all ${cellCount(block).toLocaleString("en-US")} cells for formula errors and blank cells`);
    verification.scopeNote = `${scanned.join("; ")}.`;
    if (scannedErrors > sampledErrors && !errorCells.isNullObject) {
      verification.scopeNote += ` Formula errors in: ${truncateAddressList(errorCells.address)}.`;
    }
    if (missingBlanks > acc.notWritten.length && !blankCells.isNullObject) {
      verification.scopeNote += ` Blank cells in: ${truncateAddressList(blankCells.address)}.`;
    }
  }

  const extended = tablesAfter && options.tablesBefore
    ? tablesAfter.items.filter((table) => !options.tablesBefore?.includes(table.name))
    : [];
  if (extended.length > 0) {
    const ranges = extended.map((table) => {
      const range = table.getRange();
      range.load("address");
      const header = table.getHeaderRowRange();
      header.load("values");
      return { name: table.name, range, header };
    });
    await context.sync();
    verification.tableExtensions = ranges.map(({ name, range, header }) => ({
      name,
      address: stripSheet(range.address),
      newHeaders: (header.values[0] ?? []).map(String).filter((text) => /^Column\d+$/u.test(text)),
    }));
  }

  if (canReadValidation && target.dataValidation.valid !== true) {
    await collectInvalidCells(context, sheet, block, sent, verification);
  }

  if (acc.converted.length > 0) {
    const shown = acc.converted.slice(0, LISTED_CELL_LIMIT).map((issue) => {
      const cell = sheet.getRange(issue.address);
      cell.load("text");
      return { issue, cell };
    });
    await context.sync();
    for (const { issue, cell } of shown) {
      const text = cell.text[0]?.[0];
      if (typeof text === "string") issue.shown = text;
    }
  }

  return verification;
}

function truncateAddressList(address: string): string {
  const areas = address.split(",").map(stripSheet);
  return listCells(areas, 6);
}

async function collectInvalidCells(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  block: WriteBlock,
  sent: SentGrid,
  verification: WriteVerification,
): Promise<void> {
  const probes: Array<{ address: string; range: Excel.Range }> = [];
  let candidates = 0;
  for (let r = 0; r < block.rowCount; r += 1) {
    for (let c = 0; c < block.columnCount; c += 1) {
      if (!sendsContent(sent(r, c))) continue;
      candidates += 1;
      if (probes.length >= PER_CELL_PROBE_LIMIT) continue;
      const address = cellAt(block, r, c);
      const range = sheet.getRange(address);
      range.dataValidation.load("valid");
      probes.push({ address, range });
    }
  }
  if (probes.length > 0) await context.sync();
  verification.invalidCells = probes
    .filter((probe) => probe.range.dataValidation.valid === false)
    .map((probe) => probe.address);
  const firstInvalid = verification.invalidCells[0];
  if (firstInvalid) {
    const validation = sheet.getRange(firstInvalid).dataValidation;
    validation.load("type,rule");
    await context.sync();
    verification.invalidRule = describeValidationRule(validation.type, validation.rule);
  }
  if (candidates > probes.length) {
    verification.invalidSummary =
      `data validation was checked cell by cell for the first ${probes.length} of ${candidates} cells; ` +
      "the range as a whole contains values that violate it";
  }
}

function ruleOperand(operand: string | number | Excel.Range | undefined): string {
  if (typeof operand === "string") return operand;
  if (typeof operand === "number") return String(operand);
  return "a range";
}

function describeValidationRule(type: string, rule: Excel.DataValidationRule): string {
  if (rule.list) return `list: ${ruleOperand(rule.list.source)}`;
  const basic = rule.wholeNumber ?? rule.decimal ?? rule.textLength;
  if (basic) {
    const bounds = basic.formula2 !== undefined
      ? `${ruleOperand(basic.formula1)} and ${ruleOperand(basic.formula2)}`
      : ruleOperand(basic.formula1);
    return `${type} ${basic.operator} ${bounds}`;
  }
  if (rule.custom) return `custom: ${rule.custom.formula}`;
  return type;
}

function looksLikeHeaderRename(issue: CellIssue): boolean {
  const sent = issue.sent.startsWith("\"") ? issue.sent.slice(1, -1) : issue.sent;
  const got = issue.got.startsWith("\"") ? issue.got.slice(1, -1) : issue.got;
  return got.startsWith(sent) && /^\d+$/u.test(got.slice(sent.length));
}

function issueLines(issues: CellIssue[], describe: (issue: CellIssue) => string): string[] {
  const lines = issues.slice(0, LISTED_CELL_LIMIT).map((issue) => `- ${describe(issue)}`);
  if (issues.length > LISTED_CELL_LIMIT) lines.push(`- …and ${issues.length - LISTED_CELL_LIMIT} more`);
  return lines;
}

export function formatVerificationProblems(
  verification: WriteVerification,
  qualify: (address: string) => string,
): string[] {
  const lines: string[] = [];

  if (verification.notWrittenCount > 0) {
    lines.push(`⚠️ **${verification.notWrittenCount} cell(s) did not take the value sent** (Excel accepted the write but the cell does not hold it):`);
    lines.push(...issueLines(verification.notWritten, (issue) =>
      `${qualify(issue.address)}: sent ${issue.sent}, cell holds ${issue.got}`));
    if (verification.notWritten.length < verification.notWrittenCount) {
      lines.push(`- ${verification.notWrittenCount - verification.notWritten.length} more blank cell(s) outside the rows read back.`);
    }
    lines.push("");
  }

  if (verification.converted.length > 0) {
    lines.push(`⚠️ **Excel converted ${verification.converted.length} text value(s):**`);
    lines.push(...issueLines(verification.converted, (issue) =>
      `${qualify(issue.address)}: sent ${issue.sent}, stored ${issue.got}${issue.shown !== undefined && issue.shown !== issue.got ? ` (shown as "${issue.shown}")` : ""}`));
    lines.push("To keep a value as text, rewrite it with a leading apostrophe, e.g. \"'00123\" (the apostrophe is not stored).");
    lines.push("");
  }

  if (verification.altered.length > 0) {
    lines.push(`⚠️ **Excel stored ${verification.altered.length} value(s) differently from what was sent:**`);
    lines.push(...issueLines(verification.altered, (issue) =>
      `${qualify(issue.address)}: sent ${issue.sent}, cell holds ${issue.got}`));
    if (verification.altered.some(looksLikeHeaderRename)) {
      lines.push("(Table headers must be unique; Excel renames a duplicate header by adding a number.)");
    }
    lines.push("");
  }

  if (verification.formulaErrorCount > 0) {
    lines.push(`⚠️ **${verification.formulaErrorCount} formula error(s):**`);
    for (const error of verification.formulaErrors.slice(0, LISTED_CELL_LIMIT)) {
      lines.push(`- ${qualify(error.address)}: ${error.error}${error.formula ? ` (formula: ${error.formula})` : ""}`);
    }
    if (verification.formulaErrors.length > LISTED_CELL_LIMIT) {
      lines.push(`- …and ${verification.formulaErrors.length - LISTED_CELL_LIMIT} more`);
    }
    if (verification.formulaErrors.some((error) => error.error === "#SPILL!")) {
      lines.push("#SPILL! means the formula's results need empty cells below/right of it; clear them or move the formula.");
    }
    lines.push("");
  }

  if (verification.invalidCells.length > 0 || verification.invalidSummary) {
    const count = verification.invalidCells.length;
    const rule = verification.invalidRule ? ` (${verification.invalidRule})` : "";
    lines.push(count > 0
      ? `⚠️ **${count} value(s) violate the cells' data validation${rule}:** ${listCells(verification.invalidCells.map(qualify))}.`
      : "⚠️ **Some values violate the cells' data validation.**");
    if (verification.invalidSummary) lines.push(`(${verification.invalidSummary}.)`);
    lines.push("Excel stores them anyway, but a user typing them would be refused. Write an allowed value, or tell the user.");
    lines.push("");
  }

  while (lines.at(-1) === "") lines.pop();
  return lines;
}

export function formatVerificationNotes(verification: WriteVerification): string[] {
  const lines: string[] = [];
  if (verification.manualCalculation) {
    lines.push(
      `ℹ️ Calculation is set to ${verification.manualCalculation}. ` +
      (verification.calculatedWritten ? "The written cells were calculated, so their values are current, but other formulas " : "Formulas ") +
      "that depend on them still show old results until the user recalculates (F9) or switches calculation to Automatic.",
    );
  }
  for (const table of verification.tableExtensions) {
    const headers = table.newHeaders.length > 0
      ? ` It named the new column(s) ${table.newHeaders.map((header) => `"${header}"`).join(", ")}; rename the header cell with write_cells if that is not wanted.`
      : "";
    lines.push(`ℹ️ Excel added the written cells to table **${table.name}**, which now covers ${table.address}.${headers}`);
  }
  if (verification.mode === "sampled" && verification.scopeNote) {
    lines.push(`ℹ️ Verification is partial: the range is too large to read back in one call, so I ${verification.scopeNote}`);
  }
  return lines;
}

export function hasVerificationProblems(verification: WriteVerification): boolean {
  return verification.notWrittenCount > 0 ||
    verification.converted.length > 0 ||
    verification.altered.length > 0 ||
    verification.formulaErrorCount > 0 ||
    verification.invalidCells.length > 0 ||
    verification.invalidSummary !== null;
}

export function replacedFormulas(
  block: WriteBlock,
  beforeFormulas: CellScalar[][],
  afterFormulas: CellScalar[][],
): Array<{ address: string; formula: string }> {
  const replaced: Array<{ address: string; formula: string }> = [];
  beforeFormulas.forEach((row, r) => row.forEach((before, c) => {
    const after = afterFormulas[r]?.[c];
    if (typeof before !== "string" || !before.startsWith("=")) return;
    if (after === null || after === undefined || after === "") return;
    if (typeof after === "string" && after.startsWith("=")) return;
    replaced.push({ address: cellAt(block, r, c), formula: before });
  }));
  return replaced;
}

export function formatReplacedFormulas(
  sheetName: string,
  replaced: ReadonlyArray<{ address: string; formula: string }>,
): string | null {
  if (replaced.length === 0) return null;
  const samples = replaced.slice(0, 3).map(({ address, formula }) => `${address} was \`${formula}\``);
  return `⚠️ ${replaced.length} formula(s) on sheet "${sheetName}" replaced by values: ` +
    `${listCells(replaced.map(({ address }) => address))}. ${samples.join(", ")}. ` +
    "These cells no longer update when their inputs change; if the user did not ask to hard-code them, write formulas instead.";
}

export type WriteInterruption = { kind: "stopped" } | { kind: "failed"; reason: string };

function rowSpan(block: WriteBlock, offset: number, count: number): string {
  const first = block.startRow + offset;
  return count === 1 ? `row ${first}` : `rows ${first}–${first + count - 1}`;
}

export function formatInterruptedWrite(args: {
  cause: WriteInterruption;
  block: WriteBlock;
  sheetName: string | null;
  fullAddr: string;
  writtenRows: number;
  uncertainRows: number;
}): string {
  const { cause, block, sheetName, fullAddr, writtenRows, uncertainRows } = args;
  const reason = cause.kind === "failed" ? cause.reason : "";

  if (sheetName === null || writtenRows + uncertainRows === 0) {
    return cause.kind === "stopped"
      ? `⏹ **Stopped before writing** — nothing changed in ${fullAddr}.`
      : `Error writing cells — nothing was written to ${fullAddr}. ${reason}`;
  }

  const qualify = (address: string): string => qualifiedAddress(sheetName, address);
  const writtenAddr = qualify(blockAddress(block, 0, writtenRows));
  const touchedRows = writtenRows + uncertainRows;
  const touchedAddr = qualify(blockAddress(block, 0, touchedRows));
  const remainingRows = block.rowCount - touchedRows;
  const lines: string[] = [];

  if (writtenRows > 0) {
    lines.push(cause.kind === "stopped"
      ? `⏹ **Stopped after writing ${writtenAddr}** (${rowSpan(block, 0, writtenRows)} of ${fullAddr}) — the change stands.`
      : `⚠️ **Partly written** — ${writtenAddr} (${rowSpan(block, 0, writtenRows)} of ${fullAddr}) was written and stands.`);
  } else {
    lines.push(cause.kind === "stopped"
      ? `⏹ **Stopped while writing** ${fullAddr}.`
      : `⚠️ **Outcome unknown** for ${fullAddr}.`);
  }
  if (uncertainRows > 0) {
    lines.push(
      `${rowSpan(block, writtenRows, uncertainRows)} (${qualify(blockAddress(block, writtenRows, uncertainRows))}) ` +
      "may or may not have been written: Excel did not confirm that step.",
    );
  }
  if (remainingRows > 0) {
    lines.push(`The remaining ${remainingRows} row(s) were not written.`);
  }
  if (reason) lines.push(reason);
  lines.push("");
  lines.push(`Nothing was rolled back. Read ${touchedAddr} before retrying${remainingRows > 0 ? " the rest" : ""}.`);
  return lines.join("\n");
}

export function formatStoppedAfterWrite(fullAddr: string, block: WriteBlock): string {
  return `⏹ **Stopped after writing ${fullAddr}** (${block.rowCount}×${block.columnCount}) — the change stands. ` +
    "It was not verified because Stop was pressed; read the range before relying on it.";
}

export function chunkRowsForWrite(rows: CellScalar[][], maxChars = 2_000_000): Array<{ rowOffset: number; rows: CellScalar[][] }> {
  const chunks: Array<{ rowOffset: number; rows: CellScalar[][] }> = [];
  let current: CellScalar[][] = [];
  let currentOffset = 0;
  let currentSize = 0;
  rows.forEach((row, index) => {
    const size = JSON.stringify(row).length + 1;
    if (current.length > 0 && currentSize + size > maxChars) {
      chunks.push({ rowOffset: currentOffset, rows: current });
      current = [];
      currentOffset = index;
      currentSize = 0;
    }
    current.push(row);
    currentSize += size;
  });
  if (current.length > 0) chunks.push({ rowOffset: currentOffset, rows: current });
  return chunks;
}

export function backupUnavailable(
  beforeCaptured: boolean,
  block: WriteBlock,
  fullAddr: string,
): { unavailableReason: string; unavailableNote?: string } {
  if (beforeCaptured) return { unavailableReason: CHECKPOINT_SKIPPED_REASON };
  const cells = block.rowCount * block.columnCount;
  return {
    unavailableReason: tooLargeToBackUpReason(cells),
    unavailableNote:
      `ℹ️ No backup: this change covers ${cells.toLocaleString("en-US")} cells, more than the ` +
      `${MAX_RECOVERY_CELLS.toLocaleString("en-US")} a backup can hold. ${fullAddr} was empty before, so clearing it undoes the change.`,
  };
}

export function sampledChanges(
  sheetName: string,
  block: WriteBlock,
  verification: WriteVerification,
  sentCells: number,
): WorkbookCellChangeSummary {
  const head = verification.sampleRows[0];
  const changedCount = Math.max(0, sentCells - verification.notWrittenCount);
  if (!head) return { changedCount, truncated: changedCount > 0, sample: [] };
  const empty = head.values.map((row) => row.map(() => ""));
  const summary = buildWorkbookCellChangeSummary({
    sheetName,
    startCell: `${colToLetter(block.startCol)}${block.startRow}`,
    beforeValues: empty,
    beforeFormulas: empty,
    afterValues: head.values,
    afterFormulas: head.formulas,
  });
  return { ...summary, changedCount, truncated: summary.sample.length < changedCount };
}
