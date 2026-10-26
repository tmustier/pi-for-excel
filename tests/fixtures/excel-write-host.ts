/**
 * A small in-memory Excel host for write-tool tests.
 *
 * It models the Excel behaviors the write path has to tell the truth about,
 * as observed in Excel on the web (2026-10-26):
 * - a value sent to a merged cell other than its top-left is dropped;
 * - a value sent into a spilled range turns the anchor into #SPILL!;
 * - numeric-looking text is stored as a number, "'text" as text;
 * - a new formula reads back 0 under manual calculation until calculated;
 * - data validation is not enforced on writes, only reported by `valid`;
 * - writes to locked cells on a protected sheet fail with AccessDenied;
 * - loading values/formulas for more than PAYLOAD_CELL_LIMIT cells fails the
 *   sync with ResponsePayloadSizeLimitExceeded, after writes in it applied.
 *
 * It is narrow on purpose: formulas evaluate only `=REF`, `=REF+REF…` and
 * `=REF*REF`, and only the APIs the write tools call exist.
 */

interface Cell {
  value: string | number | boolean;
  formula: string;
}

interface Coord { row: number; col: number }

interface Block { top: number; left: number; bottom: number; right: number }

const PAYLOAD_CELL_LIMIT = 50_000;

function colIndex(letters: string): number {
  let col = 0;
  for (const letter of letters.toUpperCase()) col = col * 26 + letter.charCodeAt(0) - 64;
  return col - 1;
}

export function colName(index: number): string {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function parseCoord(text: string): Coord {
  const match = /^\$?([A-Z]+)\$?(\d+)$/iu.exec(text);
  if (!match?.[1] || !match[2]) throw new Error(`Unsupported test address: ${text}`);
  return { row: Number(match[2]) - 1, col: colIndex(match[1]) };
}

function parseBlock(address: string): Block {
  const cellPart = address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address;
  const [startText = "", endText = startText] = cellPart.split(":");
  const start = parseCoord(startText);
  const end = parseCoord(endText);
  return { top: start.row, left: start.col, bottom: end.row, right: end.col };
}

function blockText(block: Block): string {
  const start = `${colName(block.left)}${block.top + 1}`;
  const end = `${colName(block.right)}${block.bottom + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function key(row: number, col: number): string {
  return `${row}:${col}`;
}

export class OfficeError extends Error {
  readonly code: string;
  readonly debugInfo: { code: string; errorLocation: string };

  constructor(code: string, message: string, errorLocation = "") {
    super(message);
    this.code = code;
    this.debugInfo = { code, errorLocation };
  }
}

interface Merge { block: Block }
interface Spill { anchor: Coord; formula: string; values: (string | number)[] }
interface ListValidation { block: Block; allowed: string[] }

export class FakeSheet {
  readonly name: string;
  readonly cells = new Map<string, Cell>();
  readonly merges: Merge[] = [];
  readonly spills: Spill[] = [];
  readonly validations: ListValidation[] = [];
  readonly unlocked = new Set<string>();
  readonly protection = { protected: false, load: (_properties?: string): void => {} };
  readonly comments = { items: [] as never[], load: (_properties?: string): void => {} };
  readonly host: FakeExcelHost;

  constructor(host: FakeExcelHost, name: string) {
    this.host = host;
    this.name = name;
  }

  load(_properties?: string | string[]): void {}

  getRange(address: string): FakeRange {
    return new FakeRange(this, parseBlock(address));
  }

  // ── Fixture setup ──

  set(address: string, value: string | number | boolean): void {
    const { row, col } = parseCoord(address);
    if (typeof value === "string" && value.startsWith("=")) {
      this.cells.set(key(row, col), { value: this.evaluate(value), formula: value });
    } else {
      this.cells.set(key(row, col), { value, formula: "" });
    }
  }

  merge(address: string): void {
    this.merges.push({ block: parseBlock(address) });
  }

  spill(anchorAddress: string, formula: string, values: (string | number)[]): void {
    const anchor = parseCoord(anchorAddress);
    this.spills.push({ anchor, formula, values });
    this.cells.set(key(anchor.row, anchor.col), { value: values[0] ?? "", formula });
    values.slice(1).forEach((value, index) => {
      this.cells.set(key(anchor.row + index + 1, anchor.col), { value, formula: "" });
    });
  }

  listValidation(address: string, allowed: string[]): void {
    this.validations.push({ block: parseBlock(address), allowed });
  }

  protect(unlockedAddresses: string[] = []): void {
    this.protection.protected = true;
    for (const address of unlockedAddresses) {
      const { row, col } = parseCoord(address);
      this.unlocked.add(key(row, col));
    }
  }

  value(address: string): string | number | boolean {
    const { row, col } = parseCoord(address);
    return this.cells.get(key(row, col))?.value ?? "";
  }

  formula(address: string): string {
    const { row, col } = parseCoord(address);
    return this.cells.get(key(row, col))?.formula ?? "";
  }

  // ── Behavior ──

  mergeAt(row: number, col: number): Merge | undefined {
    return this.merges.find(({ block }) =>
      row >= block.top && row <= block.bottom && col >= block.left && col <= block.right);
  }

  spillAt(row: number, col: number): Spill | undefined {
    return this.spills.find((spill) =>
      col === spill.anchor.col && row >= spill.anchor.row && row < spill.anchor.row + spill.values.length);
  }

  evaluate(formula: string): string | number | boolean {
    const body = formula.slice(1);
    const product = /^([A-Z]+\d+)\*([A-Z]+\d+)$/iu.exec(body);
    if (product?.[1] && product[2]) return this.numberAt(product[1]) * this.numberAt(product[2]);
    if (/^[A-Z]+\d+(\+[A-Z]+\d+)*$/iu.test(body)) {
      return body.split("+").reduce((total, ref) => total + this.numberAt(ref), 0);
    }
    return 0;
  }

  private numberAt(ref: string): number {
    const value = this.value(ref);
    return typeof value === "number" ? value : 0;
  }

  writeCell(row: number, col: number, sent: unknown): void {
    if (sent === null || sent === undefined) return;
    const merge = this.mergeAt(row, col);
    if (merge && (row !== merge.block.top || col !== merge.block.left)) return; // dropped by Excel
    const spill = this.spillAt(row, col);
    if (spill && (row !== spill.anchor.row) && sent !== "") {
      this.cells.set(key(spill.anchor.row, spill.anchor.col), { value: "#SPILL!", formula: spill.formula });
    }
    if (typeof sent === "string" && sent.startsWith("=")) {
      const value = this.host.calculationMode === "Automatic" ? this.evaluate(sent) : 0;
      this.cells.set(key(row, col), { value, formula: sent });
      return;
    }
    if (typeof sent === "string") {
      if (sent.startsWith("'")) this.cells.set(key(row, col), { value: sent.slice(1), formula: "" });
      else if (/^-?\d+(\.\d+)?$/u.test(sent)) this.cells.set(key(row, col), { value: Number(sent), formula: "" });
      else this.cells.set(key(row, col), { value: sent, formula: "" });
      return;
    }
    if (typeof sent === "number" || typeof sent === "boolean") {
      this.cells.set(key(row, col), { value: sent, formula: "" });
    }
  }

  isLocked(row: number, col: number): boolean {
    return !this.unlocked.has(key(row, col));
  }

  validationAt(row: number, col: number): ListValidation | undefined {
    return this.validations.find(({ block }) =>
      row >= block.top && row <= block.bottom && col >= block.left && col <= block.right);
  }
}

class FakeAreas {
  readonly isNullObject: boolean;
  readonly address: string;
  readonly cellCount: number;
  readonly areas: { items: Array<{ address: string; rowCount: number; columnCount: number }> };

  constructor(sheetName: string, blocks: Block[]) {
    this.isNullObject = blocks.length === 0;
    this.address = blocks.map((block) => `${sheetName}!${blockText(block)}`).join(",");
    this.cellCount = blocks.reduce((total, block) =>
      total + (block.bottom - block.top + 1) * (block.right - block.left + 1), 0);
    // Excel on the web reports a merged area by its top-left cell.
    this.areas = {
      items: blocks.map((block) => ({
        address: `${sheetName}!${colName(block.left)}${block.top + 1}`,
        rowCount: block.bottom - block.top + 1,
        columnCount: block.right - block.left + 1,
      })),
    };
  }

  load(_properties?: string): void {}
}

export class FakeRange {
  readonly sheet: FakeSheet;
  readonly block: Block;
  readonly isNullObject = false;
  readonly format: { protection: { locked: boolean | null; load: (properties?: string) => void } };
  readonly dataValidation: {
    readonly valid: boolean | null;
    readonly type: string;
    readonly rule: { list?: { source: string } };
    load: (properties?: string) => void;
  };

  constructor(sheet: FakeSheet, block: Block) {
    this.sheet = sheet;
    this.block = block;
    const cells = (): Coord[] => this.coords();
    this.format = {
      protection: {
        get locked(): boolean | null {
          const states = new Set(cells().map(({ row, col }) => sheet.isLocked(row, col)));
          return states.size === 1 ? [...states][0] ?? null : null;
        },
        load: (): void => {},
      },
    };
    this.dataValidation = {
      get valid(): boolean | null {
        const states = new Set(cells().map(({ row, col }) => {
          const rule = sheet.validationAt(row, col);
          const value = sheet.cells.get(key(row, col))?.value ?? "";
          return !rule || value === "" || rule.allowed.includes(String(value));
        }));
        return states.size === 1 ? [...states][0] ?? null : null;
      },
      get type(): string {
        return cells().some(({ row, col }) => sheet.validationAt(row, col)) ? "List" : "None";
      },
      get rule(): { list?: { source: string } } {
        const first = cells().map(({ row, col }) => sheet.validationAt(row, col)).find((rule) => rule);
        return first ? { list: { source: first.allowed.join(",") } } : {};
      },
      load: (): void => {},
    };
  }

  get address(): string {
    return `${this.sheet.name}!${blockText(this.block)}`;
  }

  get rowCount(): number {
    return this.block.bottom - this.block.top + 1;
  }

  get columnCount(): number {
    return this.block.right - this.block.left + 1;
  }

  private coords(): Coord[] {
    const out: Coord[] = [];
    for (let row = this.block.top; row <= this.block.bottom; row += 1) {
      for (let col = this.block.left; col <= this.block.right; col += 1) out.push({ row, col });
    }
    return out;
  }

  load(properties?: string | string[]): void {
    const list = Array.isArray(properties) ? properties.join(",") : properties ?? "";
    if (/\b(values|formulas|text)\b/u.test(list)) this.sheet.host.pendingLoadedCells += this.rowCount * this.columnCount;
  }

  private grid(read: (cell: Cell | undefined) => string | number | boolean): (string | number | boolean)[][] {
    const rows: (string | number | boolean)[][] = [];
    for (let row = this.block.top; row <= this.block.bottom; row += 1) {
      const values: (string | number | boolean)[] = [];
      for (let col = this.block.left; col <= this.block.right; col += 1) {
        values.push(read(this.sheet.cells.get(key(row, col))));
      }
      rows.push(values);
    }
    return rows;
  }

  get values(): (string | number | boolean)[][] {
    return this.grid((cell) => cell?.value ?? "");
  }

  set values(values: unknown[][]) {
    this.sheet.host.queueWrite(() => {
      if (this.sheet.protection.protected && this.coords().some(({ row, col }) => this.sheet.isLocked(row, col))) {
        throw new OfficeError("AccessDenied", "You cannot perform the requested operation.", "Range.values");
      }
      values.forEach((rowValues, r) => rowValues.forEach((value, c) => {
        this.sheet.writeCell(this.block.top + r, this.block.left + c, value);
      }));
    });
  }

  get formulas(): (string | number | boolean)[][] {
    return this.grid((cell) => cell?.formula ?? "");
  }

  set formulas(formulas: unknown[][]) {
    this.values = formulas;
  }

  get text(): string[][] {
    return this.grid((cell) => String(cell?.value ?? "")).map((row) => row.map(String));
  }

  get hasSpill(): boolean | null {
    const states = new Set(this.coords().map(({ row, col }) => this.sheet.spillAt(row, col) !== undefined));
    return states.size === 1 ? [...states][0] ?? null : null;
  }

  getCell(row: number, col: number): FakeRange {
    const top = this.block.top + row;
    const left = this.block.left + col;
    return new FakeRange(this.sheet, { top, left, bottom: top, right: left });
  }

  getRow(row: number): FakeRange {
    const top = this.block.top + row;
    return new FakeRange(this.sheet, { ...this.block, top, bottom: top });
  }

  autoFill(destination: FakeRange, _type?: string): void {
    this.sheet.host.queueWrite(() => {
      const source = this.sheet.formula(blockText({ ...this.block, bottom: this.block.top, right: this.block.left }));
      for (const { row, col } of destination.coords()) {
        const dr = row - this.block.top;
        const dc = col - this.block.left;
        const shifted = source.replace(/([A-Z]+)(\d+)/gu, (_match, letters: string, digits: string) =>
          `${colName(colIndex(letters) + dc)}${Number(digits) + dr}`);
        this.sheet.writeCell(row, col, shifted);
      }
    });
  }

  calculate(): void {
    this.sheet.host.queueWrite(() => {
      for (const { row, col } of this.coords()) {
        const cell = this.sheet.cells.get(key(row, col));
        if (cell?.formula) cell.value = this.sheet.evaluate(cell.formula);
      }
    });
  }

  getMergedAreasOrNullObject(): FakeAreas {
    const blocks = this.sheet.merges
      .filter(({ block }) =>
        block.top >= this.block.top && block.top <= this.block.bottom &&
        block.left >= this.block.left && block.left <= this.block.right)
      .map(({ block }) => block);
    return new FakeAreas(this.sheet.name, blocks);
  }

  getSpillParentOrNullObject(): FakeRange | { isNullObject: true; load: () => void } {
    const spill = this.rowCount === 1 && this.columnCount === 1
      ? this.sheet.spillAt(this.block.top, this.block.left)
      : undefined;
    if (!spill) return { isNullObject: true, load: (): void => {} };
    return new FakeRange(this.sheet, {
      top: spill.anchor.row, left: spill.anchor.col, bottom: spill.anchor.row, right: spill.anchor.col,
    });
  }

  getUsedRangeOrNullObject(_valuesOnly?: boolean): FakeAreas {
    const used = this.coords().filter(({ row, col }) => {
      const cell = this.sheet.cells.get(key(row, col));
      return cell !== undefined && (cell.formula !== "" || cell.value !== "");
    });
    if (used.length === 0) return new FakeAreas(this.sheet.name, []);
    const block = {
      top: Math.min(...used.map(({ row }) => row)),
      bottom: Math.max(...used.map(({ row }) => row)),
      left: Math.min(...used.map(({ col }) => col)),
      right: Math.max(...used.map(({ col }) => col)),
    };
    return new FakeAreas(this.sheet.name, [block]);
  }

  getSpecialCellsOrNullObject(type: string, valueType?: string): FakeAreas {
    const matches = this.coords().filter(({ row, col }) => {
      const cell = this.sheet.cells.get(key(row, col));
      if (type === "Blanks") return cell === undefined || (cell.formula === "" && cell.value === "");
      if (type === "Formulas" && valueType === "Errors") {
        return cell !== undefined && cell.formula !== "" && typeof cell.value === "string" && cell.value.startsWith("#");
      }
      return false;
    });
    return new FakeAreas(this.sheet.name, matches.map(({ row, col }) => ({ top: row, bottom: row, left: col, right: col })));
  }
}

export class FakeExcelHost {
  calculationMode = "Automatic";
  pendingLoadedCells = 0;
  private readonly sheets = new Map<string, FakeSheet>();
  private queued: Array<() => void> = [];
  private active: FakeSheet;
  /** Every sync, for asserting how many round trips a tool made. */
  syncCount = 0;
  /** Fail the sync that would apply the Nth queued write (1-based) with this error; the write is not applied. */
  failWrite: { at: number; error: Error } | null = null;
  private writesApplied = 0;

  constructor(sheetNames: string[] = ["Sheet1"]) {
    for (const name of sheetNames) this.sheets.set(name, new FakeSheet(this, name));
    const first = this.sheets.get(sheetNames[0] ?? "Sheet1");
    if (!first) throw new Error("FakeExcelHost needs a sheet");
    this.active = first;
  }

  sheet(name: string): FakeSheet {
    const sheet = this.sheets.get(name);
    if (!sheet) throw new Error(`No fake sheet ${name}`);
    return sheet;
  }

  queueWrite(write: () => void): void {
    this.queued.push(write);
  }

  readonly context = {
    workbook: {
      worksheets: {
        getActiveWorksheet: (): FakeSheet => this.active,
        getItem: (name: string): FakeSheet => {
          const sheet = this.sheets.get(name);
          if (!sheet) {
            // Office.js fails at sync; failing at the call is close enough for these tests.
            throw new OfficeError("ItemNotFound", "The requested resource doesn't exist.", "WorksheetCollection.getItem");
          }
          return sheet;
        },
        load: (_properties?: string): void => {},
        items: [] as FakeSheet[],
      },
      application: {
        calculationMode: "Automatic",
        load: (_properties?: string): void => {},
      },
      protection: { protected: false, load: (_properties?: string): void => {} },
    },
    sync: (): Promise<void> => {
      this.syncCount += 1;
      this.context.workbook.worksheets.items = [...this.sheets.values()];
      this.context.workbook.application.calculationMode = this.calculationMode;
      const writes = this.queued;
      this.queued = [];
      const loaded = this.pendingLoadedCells;
      this.pendingLoadedCells = 0;
      try {
        for (const write of writes) {
          if (this.failWrite && this.writesApplied + 1 === this.failWrite.at) throw this.failWrite.error;
          write();
          this.writesApplied += 1;
        }
      } catch (error) {
        return Promise.reject(error instanceof Error ? error : new Error(String(error)));
      }
      if (loaded > PAYLOAD_CELL_LIMIT) {
        return Promise.reject(new OfficeError(
          "ResponsePayloadSizeLimitExceeded",
          "The response payload size has exceeded the limit.",
        ));
      }
      return Promise.resolve();
    },
  };
}

/** Install the fake as the global `Excel` for the duration of `action`. */
export async function withFakeExcel<T>(fake: FakeExcelHost, action: () => Promise<T>): Promise<T> {
  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (context: unknown) => Promise<TResult>): Promise<TResult> => callback(fake.context),
  });
  try {
    return await action();
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}
