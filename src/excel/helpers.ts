export interface RangeRef {
  sheet?: string;
  address: string;
}

const CELL_EDIT_MODE_ERROR_CODE = "InvalidOperationInCellEditMode";

export const CELL_EDIT_WAIT_MS = 60_000;
const CELL_EDIT_RETRY_FIRST_DELAY_MS = 250;
const CELL_EDIT_RETRY_MAX_DELAY_MS = 2_000;

export class ExcelCellEditModeError extends Error {
  readonly code = "excel_cell_edit_mode";
  readonly waitedMs: number;
  readonly earlierStepReachedExcel: boolean;

  constructor(waitedMs: number, earlierStepReachedExcel: boolean) {
    const waited = waitedMs > 0 ? ` Waited ${Math.round(waitedMs / 1000)} s for the user to finish.` : "";
    const applied = earlierStepReachedExcel
      ? " An earlier step of this operation had already reached Excel, so it was not retried automatically (a retry could apply a change twice); check the affected range before retrying."
      : " Nothing was read or changed.";
    super(
      `Excel is in cell-edit mode: the user is typing in a cell, and Excel rejects every workbook request until they press Enter, Tab or Esc.${waited}${applied}` +
      " Ask the user to finish editing the cell, then retry.",
    );
    this.name = "ExcelCellEditModeError";
    this.waitedMs = waitedMs;
    this.earlierStepReachedExcel = earlierStepReachedExcel;
  }
}

type CellEditWaitListener = (waiting: boolean) => void;
const cellEditWaitListeners = new Set<CellEditWaitListener>();
let cellEditWaiters = 0;

export function onCellEditWait(listener: CellEditWaitListener): () => void {
  cellEditWaitListeners.add(listener);
  return () => cellEditWaitListeners.delete(listener);
}

function setWaiting(delta: 1 | -1): void {
  const before = cellEditWaiters;
  cellEditWaiters = Math.max(0, cellEditWaiters + delta);
  if ((before === 0) === (cellEditWaiters === 0)) return;
  for (const listener of cellEditWaitListeners) listener(cellEditWaiters > 0);
}

export const EXCEL_RUN_TIMEOUT_MS = 60_000;

export class ExcelTimeoutError extends Error {
  readonly code = "excel_timeout";
  readonly outcomeUnknown = true;
  readonly timeoutMs: number;
  readonly consecutiveTimeouts: number;

  constructor(timeoutMs: number, consecutiveTimeouts: number) {
    const repeated = consecutiveTimeouts > 1
      ? ` Excel has not answered ${consecutiveTimeouts} requests in a row; stop calling workbook tools and ask the user to reload the Excel page (Pi keeps its chats), then retry.`
      : " If this happens again, ask the user to reload the Excel page.";
    super(
      `Excel did not answer within ${Math.round(timeoutMs / 1000)} s. The outcome is unknown: the request may or may not have been applied.` +
      ` Read the affected range back before retrying a change.${repeated}`,
    );
    this.name = "ExcelTimeoutError";
    this.timeoutMs = timeoutMs;
    this.consecutiveTimeouts = consecutiveTimeouts;
  }
}

export class ExcelWaitStoppedError extends Error {
  readonly code = "excel_wait_stopped";

  constructor() {
    super("Stopped by the user while waiting for them to finish editing a cell. Nothing was read or changed.");
    this.name = "ExcelWaitStoppedError";
  }
}

let consecutiveTimeouts = 0;
let stopGeneration = 0;
const cellEditSleepers = new Set<() => void>();

export function cancelCellEditWaits(): void {
  stopGeneration += 1;
  for (const wake of cellEditSleepers) wake();
}

function sleepUnlessCancelled(ms: number): Promise<void> {
  return new Promise((resolve) => {
    const wake = () => {
      clearTimeout(timer);
      cellEditSleepers.delete(wake);
      resolve();
    };
    const timer = setTimeout(wake, ms);
    cellEditSleepers.add(wake);
  });
}

export interface ExcelRunOptions {
  waitForCellEdit?: boolean;
  cellEditWaitMs?: number;
  timeoutMs?: number;
}

function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => {
      consecutiveTimeouts += 1;
      reject(new ExcelTimeoutError(timeoutMs, consecutiveTimeouts));
    }, timeoutMs);
  });
  promise.catch(() => {});
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

/**
 * Run an Office.js Excel operation.
 *
 * - **Timeout:** each `Excel.run` is raced against a timeout. When Excel does
 *   not answer, it throws {@link ExcelTimeoutError} ("outcome unknown"), which
 *   ends the tool call and so releases the workbook lock. The batch itself
 *   cannot be cancelled and may still complete.
 * - **Cell-edit mode:** while the user is typing in a cell, Excel rejects every
 *   batch with `InvalidOperationInCellEditMode`. The batch is retried with
 *   backoff until the user leaves edit mode, but only when no earlier
 *   `context.sync()` of this run had succeeded: re-running a half-applied
 *   mutation could apply it twice. When the wait runs out, it throws
 *   {@link ExcelCellEditModeError}; when Stop ends the wait,
 *   {@link ExcelWaitStoppedError}.
 */
export async function excelRun<T>(
  fn: (context: Excel.RequestContext) => Promise<T>,
  options: ExcelRunOptions = {},
): Promise<T> {
  const waitForCellEdit = options.waitForCellEdit ?? true;
  const waitMs = options.cellEditWaitMs ?? CELL_EDIT_WAIT_MS;
  const timeoutMs = options.timeoutMs ?? EXCEL_RUN_TIMEOUT_MS;
  const startedAt = Date.now();
  const generation = stopGeneration;
  let delay = CELL_EDIT_RETRY_FIRST_DELAY_MS;
  let waiting = false;

  try {
    for (;;) {
      let committedBatches = 0;
      try {
        const result = await withTimeout(Excel.run(async (context) => {
          const sync = context.sync.bind(context);
          context.sync = async <V>(passThroughValue?: V): Promise<V> => {
            const value = await sync(passThroughValue);
            committedBatches += 1;
            return value;
          };
          return fn(context);
        }), timeoutMs);
        consecutiveTimeouts = 0;
        return result;
      } catch (error) {
        const inCellEditMode = typeof error === "object" && error !== null
          && "code" in error && error.code === CELL_EDIT_MODE_ERROR_CODE;
        if (!inCellEditMode) throw error;
        consecutiveTimeouts = 0;
        const waitedMs = Date.now() - startedAt;
        if (!waitForCellEdit || committedBatches > 0 || waitedMs + delay > waitMs) {
          throw new ExcelCellEditModeError(waitedMs, committedBatches > 0);
        }
        if (!waiting) {
          waiting = true;
          setWaiting(1);
        }
        await sleepUnlessCancelled(delay);
        if (stopGeneration !== generation) throw new ExcelWaitStoppedError();
        delay = Math.min(CELL_EDIT_RETRY_MAX_DELAY_MS, Math.round(delay * 1.6));
      }
    }
  } finally {
    if (waiting) setWaiting(-1);
  }
}

export function parseRangeRef(ref: string): RangeRef {
  if (ref.includes("!")) {
    const idx = ref.lastIndexOf("!");
    const sheet = ref
      .substring(0, idx)
      .replace(/^'|'$/g, "")
      .replace(/''/g, "'");
    return { sheet, address: ref.substring(idx + 1) };
  }
  return { address: ref };
}

export function getRange(context: Excel.RequestContext, ref: string): { sheet: Excel.Worksheet; range: Excel.Range } {
  const parsed = parseRangeRef(ref);
  const sheet = parsed.sheet
    ? context.workbook.worksheets.getItem(parsed.sheet)
    : context.workbook.worksheets.getActiveWorksheet();
  return { sheet, range: sheet.getRange(parsed.address) };
}

export type SheetSource = "explicit" | "active";

export interface ResolvedSheetRef {
  sheet: Excel.Worksheet;
  sheetName: string;
  address: string;
  sheetSource: SheetSource;
}

export async function resolveSheetRef(context: Excel.RequestContext, ref: string): Promise<ResolvedSheetRef> {
  const parsed = parseRangeRef(ref.trim());
  if (parsed.sheet === undefined) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    return { sheet, sheetName: sheet.name, address: parsed.address, sheetSource: "active" };
  }

  const sheet = context.workbook.worksheets.getItemOrNullObject(parsed.sheet);
  sheet.load("name");
  await context.sync();
  if (sheet.isNullObject) {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const names = sheets.items.map((item) => `"${item.name}"`).join(", ");
    throw new Error(`Sheet "${parsed.sheet}" not found. Sheets in this workbook: ${names}.`);
  }
  return { sheet, sheetName: sheet.name, address: parsed.address, sheetSource: "explicit" };
}

export interface CellBlock {
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
}

export function blockAddress(block: CellBlock): string {
  const start = cellAddress(block.columnIndex, block.rowIndex + 1);
  if (block.rowCount === 1 && block.columnCount === 1) return start;
  const end = cellAddress(block.columnIndex + block.columnCount - 1, block.rowIndex + block.rowCount);
  return `${start}:${end}`;
}

export function intersectBlocks(a: CellBlock, b: CellBlock): CellBlock | null {
  const rowStart = Math.max(a.rowIndex, b.rowIndex);
  const colStart = Math.max(a.columnIndex, b.columnIndex);
  const rowEnd = Math.min(a.rowIndex + a.rowCount, b.rowIndex + b.rowCount);
  const colEnd = Math.min(a.columnIndex + a.columnCount, b.columnIndex + b.columnCount);
  if (rowEnd <= rowStart || colEnd <= colStart) return null;
  return { rowIndex: rowStart, columnIndex: colStart, rowCount: rowEnd - rowStart, columnCount: colEnd - colStart };
}

export function isResponsePayloadLimitMessage(message: string): boolean {
  return /payload size has exceeded|ResponsePayloadSizeLimitExceeded/iu.test(message);
}

export function qualifiedAddress(sheetName: string, address: string): string {
  const clean = address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address;
  const escaped = sheetName.replace(/'/g, "''");
  const needsQuote = /[\s'!]/.test(sheetName);
  const quoted = needsQuote ? `'${escaped}'` : sheetName;
  return `${quoted}!${clean}`;
}

export function colToLetter(col: number): string {
  let letter = "";
  let c = col;
  while (c >= 0) {
    letter = String.fromCharCode((c % 26) + 65) + letter;
    c = Math.floor(c / 26) - 1;
  }
  return letter;
}

export function letterToCol(letters: string): number {
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return col - 1;
}

export function parseCell(cell: string): { col: number; row: number } {
  const clean = cell.includes("!") ? cell.slice(cell.indexOf("!") + 1) : cell;
  const match = clean.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) throw new Error(`Invalid cell address: ${cell}`);

  const colLetters = match[1];
  const rowDigits = match[2];
  if (colLetters === undefined || rowDigits === undefined) {
    throw new Error(`Invalid cell address: ${cell}`);
  }

  return { col: letterToCol(colLetters.toUpperCase()), row: parseInt(rowDigits, 10) };
}

export function cellAddress(col: number, row: number): string {
  return `${colToLetter(col)}${row}`;
}

export function isCellInRange(cellAddress: string, rangeAddress: string): boolean {
  const bangIndex = rangeAddress.indexOf("!");
  const cleanRange = bangIndex >= 0 ? rangeAddress.slice(bangIndex + 1) : rangeAddress;
  const colonIndex = cleanRange.indexOf(":");
  const startAddress = colonIndex >= 0 ? cleanRange.slice(0, colonIndex) : cleanRange;
  const endAddress = colonIndex >= 0 ? cleanRange.slice(colonIndex + 1) : cleanRange;
  const start = parseCell(startAddress);
  const end = parseCell(endAddress);
  const cell = parseCell(cellAddress);
  return (
    cell.col >= start.col &&
    cell.col <= end.col &&
    cell.row >= start.row &&
    cell.row <= end.row
  );
}

export function computeRangeAddress(startCell: string, rows: number, cols: number): string {
  const { col, row } = parseCell(startCell);
  return `${startCell}:${colToLetter(col + cols - 1)}${row + rows - 1}`;
}

export function cellAtOffset(rangeStart: string, rowOffset: number, colOffset: number): string {
  const { col, row } = parseCell(rangeStart);
  return cellAddress(col + colOffset, row + rowOffset);
}

function isOfficeItemNotFound(error: unknown): boolean {
  if (typeof error !== "object" || error === null) return false;

  if ("code" in error && typeof error.code === "string") {
    return error.code === "ItemNotFound";
  }

  if ("message" in error && typeof error.message === "string") {
    return /item\s*not\s*found/iu.test(error.message);
  }

  return false;
}

export async function getDirectPrecedentsSafe(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<string[][] | null> {
  try {
    const precedents = range.getDirectPrecedents();
    precedents.load("addresses");
    await context.sync();
    return precedents.addresses
      .map((s) => s.split(",").map((x) => x.trim()).filter(Boolean));
  } catch (error) {
    if (isOfficeItemNotFound(error)) {
      return [];
    }
    return null;
  }
}

export async function getDirectDependentsSafe(
  context: Excel.RequestContext,
  range: Excel.Range,
): Promise<string[][] | null> {
  try {
    const dependents = range.getDirectDependents();
    dependents.load("addresses");
    await context.sync();
    return dependents.addresses
      .map((s) => s.split(",").map((x) => x.trim()).filter(Boolean));
  } catch (error) {
    if (isOfficeItemNotFound(error)) {
      return [];
    }
    return null;
  }
}

export function padValues(values: unknown[][]): { padded: unknown[][]; rows: number; cols: number } {
  const rows = values.length;
  const cols = Math.max(...values.map((r) => r.length));
  const padded = values.map((row) => {
    const r = [...row];
    while (r.length < cols) r.push("");
    return r;
  });
  return { padded, rows, cols };
}
