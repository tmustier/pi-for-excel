import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";
import { Type } from "@sinclair/typebox";

import { colToLetter, parseCell } from "../src/excel/helpers.ts";
import { WpsHost } from "../src/host/wps-host.ts";
import type {
  WpsCountedCollection,
  WpsEtApplication,
  WpsEtRange,
  WpsEtWorkbook,
  WpsEtWorksheet,
} from "../src/host/wps/jsapi.ts";
import { createExecuteWpsJsTool } from "../src/tools/execute-wps-js.ts";
import { selectCoreToolForHost, type AnyHostSelectableTool } from "../src/tools/host-selection.ts";
import { UnsupportedHostToolError } from "../src/tools/unsupported-host-tool.ts";
import type { ReadRangeCsvDetails, WriteCellsDetails } from "../src/tools/tool-details.ts";

interface CellPoint {
  row: number;
  col: number;
}

interface Rect {
  start: CellPoint;
  end: CellPoint;
}

function parseRangeAddress(address: string): Rect {
  const clean = address.replace(/\$/gu, "");
  const parts = clean.includes(":") ? clean.split(":") : [clean, clean];
  if (parts.length !== 2) throw new Error(`Invalid fake range address: ${address}`);
  return { start: parseCell(parts[0]), end: parseCell(parts[1]) };
}

function rowCount(rect: Rect): number {
  return Math.abs(rect.end.row - rect.start.row) + 1;
}

function colCount(rect: Rect): number {
  return Math.abs(rect.end.col - rect.start.col) + 1;
}

function makeGrid(rows: number, cols: number, fill: DynamicValue): DynamicValue[][] {
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => fill));
}

function isGrid(value: DynamicValue): value is DynamicValue[][] {
  return Array.isArray(value) && value.every((row) => Array.isArray(row));
}

class FakeRange implements WpsEtRange {
  readonly Address: string;
  readonly Rows: { Count: number };
  readonly Columns: { Count: number };
  private readonly sheet: FakeWorksheet;
  private readonly rect: Rect;

  constructor(sheet: FakeWorksheet, rect: Rect) {
    this.sheet = sheet;
    this.rect = rect;
    this.Address = `${colToLetter(rect.start.col)}${rect.start.row}:${colToLetter(rect.end.col)}${rect.end.row}`;
    this.Rows = { Count: rowCount(rect) };
    this.Columns = { Count: colCount(rect) };
  }

  get Value2(): DynamicValue[][] {
    return this.sheet.readGrid(this.rect, "values");
  }

  set Value2(values: DynamicValue) {
    this.sheet.writeValues(this.rect, values);
  }

  get Formula(): DynamicValue[][] {
    return this.sheet.readGrid(this.rect, "formulas");
  }

  set Formula(values: DynamicValue) {
    this.sheet.writeFormulas(this.rect, values);
  }

  get NumberFormat(): DynamicValue[][] {
    return this.sheet.readGrid(this.rect, "formats");
  }

  Value(_rangeValueDataType?: DynamicValue, value?: DynamicValue): DynamicValue {
    if (arguments.length >= 2) {
      this.Value2 = value;
      return undefined;
    }
    return this.Value2;
  }
}

class FakeWorksheet implements WpsEtWorksheet {
  readonly Name: string;
  readonly Visible: DynamicValue;

  private readonly values: DynamicValue[][];
  private readonly formulas: DynamicValue[][];
  private readonly formats: DynamicValue[][];
  private usedAddress: string;

  constructor(args: {
    name: string;
    values: DynamicValue[][];
    formulas?: DynamicValue[][];
    formats?: DynamicValue[][];
    visible?: DynamicValue;
    usedAddress?: string;
  }) {
    this.Name = args.name;
    this.Visible = args.visible ?? true;
    this.values = args.values.map((row) => [...row]);
    this.formulas = args.formulas?.map((row) => [...row]) ?? makeGrid(args.values.length, args.values[0]?.length ?? 1, "");
    this.formats = args.formats?.map((row) => [...row]) ?? makeGrid(args.values.length, args.values[0]?.length ?? 1, "General");
    this.usedAddress = args.usedAddress ?? `A1:${colToLetter((args.values[0]?.length ?? 1) - 1)}${args.values.length}`;
  }

  get UsedRange(): WpsEtRange {
    return this.Range(this.usedAddress);
  }

  Range(address: string): WpsEtRange {
    return new FakeRange(this, parseRangeAddress(address));
  }

  readGrid(rect: Rect, kind: "values" | "formulas" | "formats"): DynamicValue[][] {
    const source = kind === "values" ? this.values : kind === "formulas" ? this.formulas : this.formats;
    const rows: DynamicValue[][] = [];
    for (let row = rect.start.row; row <= rect.end.row; row += 1) {
      const renderedRow: DynamicValue[] = [];
      for (let col = rect.start.col; col <= rect.end.col; col += 1) {
        renderedRow.push(source[row - 1]?.[col] ?? (kind === "formats" ? "General" : ""));
      }
      rows.push(renderedRow);
    }
    return rows;
  }

  writeValues(rect: Rect, rawValues: DynamicValue): void {
    const values = isGrid(rawValues) ? rawValues : [[rawValues]];
    this.writeGrid(rect, values, false);
  }

  writeFormulas(rect: Rect, rawValues: DynamicValue): void {
    const values = isGrid(rawValues) ? rawValues : [[rawValues]];
    this.writeGrid(rect, values, true);
  }

  private writeGrid(rect: Rect, values: DynamicValue[][], formulasMayBePresent: boolean): void {
    for (let row = 0; row < rowCount(rect); row += 1) {
      for (let col = 0; col < colCount(rect); col += 1) {
        const sheetRow = rect.start.row - 1 + row;
        const sheetCol = rect.start.col + col;
        while (this.values.length <= sheetRow) this.values.push([]);
        while (this.formulas.length <= sheetRow) this.formulas.push([]);
        while (this.formats.length <= sheetRow) this.formats.push([]);

        const value = values[row]?.[col] ?? "";
        this.values[sheetRow][sheetCol] = value;
        this.formats[sheetRow][sheetCol] = this.formats[sheetRow][sheetCol] ?? "General";
        this.formulas[sheetRow][sheetCol] = formulasMayBePresent && typeof value === "string" && value.startsWith("=")
          ? value
          : "";
      }
    }

    this.usedAddress = `A1:${colToLetter(Math.max(this.maxCols() - 1, rect.end.col))}${Math.max(this.values.length, rect.end.row)}`;
  }

  private maxCols(): number {
    return this.values.reduce((max, row) => Math.max(max, row.length), 0);
  }
}

class FakeWorksheetCollection implements WpsCountedCollection {
  readonly Count: number;
  private readonly sheets: FakeWorksheet[];

  constructor(sheets: FakeWorksheet[]) {
    this.sheets = sheets;
    this.Count = sheets.length;
  }

  Item(key: string | number): WpsEtWorksheet {
    if (typeof key === "number") {
      const sheet = this.sheets[key - 1];
      if (!sheet) throw new Error(`Fake sheet index not found: ${key}`);
      return sheet;
    }

    const sheet = this.sheets.find((candidate) => candidate.Name === key);
    if (!sheet) throw new Error(`Fake sheet not found: ${key}`);
    return sheet;
  }
}

class FakeWorkbook implements WpsEtWorkbook {
  readonly Name: string;
  readonly FullName: string;
  readonly Worksheets: FakeWorksheetCollection;
  readonly Sheets: FakeWorksheetCollection;

  constructor(args: { name: string; fullName: string; sheets: FakeWorksheet[] }) {
    this.Name = args.name;
    this.FullName = args.fullName;
    this.Worksheets = new FakeWorksheetCollection(args.sheets);
    this.Sheets = this.Worksheets;
  }
}

function createFakeWpsApplication(): WpsEtApplication {
  const sheet1 = new FakeWorksheet({
    name: "Sheet1",
    values: [
      ["Region", "Revenue", "Margin"],
      ["North", 100, 0.25],
      ["South", 200, 0.35],
    ],
    formulas: [
      ["", "", ""],
      ["", "", "=B2/400"],
      ["", "", "=B3/571"],
    ],
    formats: [
      ["General", "£#,##0", "0.0%"],
      ["General", "£#,##0", "0.0%"],
      ["General", "£#,##0", "0.0%"],
    ],
  });
  const hidden = new FakeWorksheet({
    name: "HiddenData",
    values: [["secret"]],
    visible: false,
  });
  const workbook = new FakeWorkbook({
    name: "Budget.xlsx",
    fullName: "C:\\Users\\alice\\秘密\\Budget.xlsx",
    sheets: [sheet1, hidden],
  });

  return {
    ActiveWorkbook: workbook,
    ActiveSheet: sheet1,
    Selection: sheet1.Range("B2:C3"),
    Worksheets: workbook.Worksheets,
    Sheets: workbook.Sheets,
  };
}

function installFakeWps(app: WpsEtApplication): () => void {
  const hadApplication = Reflect.has(globalThis, "Application");
  const previousApplication: DynamicValue = Reflect.get(globalThis, "Application");
  const hadWps = Reflect.has(globalThis, "wps");
  const previousWps: DynamicValue = Reflect.get(globalThis, "wps");

  Reflect.set(globalThis, "Application", app);
  Reflect.set(globalThis, "wps", { EtApplication: () => app });

  return () => {
    if (hadApplication) Reflect.set(globalThis, "Application", previousApplication);
    else Reflect.deleteProperty(globalThis, "Application");

    if (hadWps) Reflect.set(globalThis, "wps", previousWps);
    else Reflect.deleteProperty(globalThis, "wps");
  };
}

function createFakeTool(name: string): AnyHostSelectableTool {
  return {
    name,
    label: `${name} label`,
    description: `${name} description`,
    parameters: Type.Object({}),
    execute: () => Promise.resolve({
      content: [{ type: "text", text: `${name} office` }],
      details: undefined,
    }),
  };
}

function firstText<TDetails>(result: AgentToolResult<TDetails>): string {
  const block = result.content[0];
  if (!block || block.type !== "text") {
    throw new Error("Expected first content block to be text.");
  }
  return block.text;
}

void test("WpsHost hashes ActiveWorkbook.FullName without exposing the raw path", async () => {
  const app = createFakeWpsApplication();
  const restore = installFakeWps(app);
  try {
    const context = await new WpsHost().getWorkbookContext();
    assert.equal(context.workbookName, "Budget.xlsx");
    assert.equal(context.source, "wps.activeWorkbook.fullName");
    assert.match(context.workbookId ?? "", /^wps_path_sha256:/u);
    const serialized = JSON.stringify(context);
    assert.equal(serialized.includes("C:\\Users\\alice"), false);
    assert.equal(serialized.includes("秘密"), false);
  } finally {
    restore();
  }
});

void test("WPS get_workbook_overview reports sheets, active context, and explicit omissions", async () => {
  const app = createFakeWpsApplication();
  const restore = installFakeWps(app);
  try {
    const tool = selectCoreToolForHost("get_workbook_overview", createFakeTool("get_workbook_overview"), "wps");
    const result = await tool.execute("tool-call-1", {});
    const text = firstText(result);
    assert.match(text, /## Workbook: Budget\.xlsx/u);
    assert.match(text, /\*\*Sheet1\*\* — 3 rows × 3 cols/u);
    assert.match(text, /\*\*HiddenData\*\* \(Hidden\)/u);
    assert.match(text, /Active sheet: \*\*Sheet1\*\*/u);
    assert.match(text, /Selection: B2:C3/u);
    assert.match(text, /tables, named ranges, charts, PivotTables, shapes/u);
  } finally {
    restore();
  }
});

void test("WPS read_range supports compact, CSV details, and detailed formula/format output", async () => {
  const app = createFakeWpsApplication();
  const restore = installFakeWps(app);
  try {
    const officeTool = createFakeTool("read_range");
    const tool = selectCoreToolForHost("read_range", officeTool, "wps");

    const compact = await tool.execute("tool-call-compact", { range: "Sheet1!A1:C3" });
    assert.match(firstText(compact), /\*\*Sheet1!A1:C3\*\* \(3×3\)/u);
    assert.match(firstText(compact), /North/u);
    assert.match(firstText(compact), /Formulas:/u);

    const csv = await tool.execute("tool-call-csv", { range: "Sheet1!A1:B2", mode: "csv" });
    const csvDetails = csv.details as ReadRangeCsvDetails | undefined;
    assert.equal(csvDetails?.kind, "read_range_csv");
    assert.equal(csvDetails.startCol, 0);
    assert.equal(csvDetails.startRow, 1);
    assert.match(firstText(csv), /Region,Revenue/u);

    const detailed = await tool.execute("tool-call-detailed", { range: "Sheet1!B2:C3", mode: "detailed" });
    const detailedText = firstText(detailed);
    assert.match(detailedText, /### Formulas/u);
    assert.match(detailedText, /=B2\/400/u);
    assert.match(detailedText, /### Number Formats/u);
    assert.match(detailedText, /`0\.0%`/u);
  } finally {
    restore();
  }
});

void test("WPS write_cells blocks overwrites, verifies writes, and reports no WPS backups", async () => {
  const app = createFakeWpsApplication();
  const restore = installFakeWps(app);
  try {
    const officeTool = createFakeTool("write_cells");
    const tool = selectCoreToolForHost("write_cells", officeTool, "wps");

    const blocked = await tool.execute("tool-call-blocked", {
      start_cell: "Sheet1!A1",
      values: [["Overwrite"]],
    });
    const blockedDetails = blocked.details as WriteCellsDetails | undefined;
    assert.equal(blockedDetails?.blocked, true);
    assert.equal(blockedDetails.existingCount, 1);
    assert.match(firstText(blocked), /Write blocked/u);
    assert.match(firstText(blocked), /WPS workbook backups are not implemented/u);

    const written = await tool.execute("tool-call-write", {
      start_cell: "Sheet1!D1",
      values: [["Total"], ["=B2+C2"]],
      allow_overwrite: false,
    });
    const writeDetails = written.details as WriteCellsDetails | undefined;
    assert.equal(writeDetails?.blocked, false);
    assert.equal(writeDetails?.recovery?.status, "not_available");
    assert.match(firstText(written), /Written to \*\*Sheet1!D1:D2\*\*/u);
    assert.match(firstText(written), /Verified values/u);
    assert.match(firstText(written), /WPS workbook backups are not implemented/u);

    const readBack = await selectCoreToolForHost("read_range", createFakeTool("read_range"), "wps")
      .execute("tool-call-readback", { range: "Sheet1!D1:D2", mode: "detailed" });
    const readBackText = firstText(readBack);
    assert.match(readBackText, /Total/u);
    assert.match(readBackText, /=B2\+C2/u);
  } finally {
    restore();
  }
});

void test("execute_wps_js serializes success and errors", async () => {
  const okTool = createExecuteWpsJsTool({
    runCode: () => Promise.resolve({ ok: true, workbook: "Budget.xlsx" }),
  });
  const ok = await okTool.execute("tool-call-ok", {
    explanation: "Inspect workbook",
    code: "return { ok: true };",
  });
  assert.match(firstText(ok), /Executed WPS JSAPI: Inspect workbook/u);
  assert.match(firstText(ok), /"workbook": "Budget\.xlsx"/u);

  const errorTool = createExecuteWpsJsTool({
    runCode: () => Promise.reject(new Error("boom")),
  });
  const failed = await errorTool.execute("tool-call-error", {
    explanation: "Inspect workbook",
    code: "return Application.ActiveWorkbook.Name;",
  });
  assert.match(firstText(failed), /Error executing WPS JSAPI: boom/u);
});

void test("WPS host wiring keeps core metadata stable and registers execute_wps_js only on WPS", async () => {
  const readOfficeTool = createFakeTool("read_range");
  const readWpsTool = selectCoreToolForHost("read_range", readOfficeTool, "wps");
  assert.notEqual(readWpsTool.execute, readOfficeTool.execute);
  assert.equal(readWpsTool.name, readOfficeTool.name);
  assert.equal(readWpsTool.label, readOfficeTool.label);
  assert.equal(readWpsTool.description, readOfficeTool.description);
  assert.equal(readWpsTool.parameters, readOfficeTool.parameters);

  const fillFormulaLikeTool = { ...readOfficeTool, name: "fill_formula" };
  const unsupported = selectCoreToolForHost("fill_formula", fillFormulaLikeTool, "wps");
  await assert.rejects(
    async () => unsupported.execute("tool-call-unsupported", {}),
    (error: DynamicValue) => {
      assert.ok(error instanceof UnsupportedHostToolError);
      assert.equal(error.toolName, "fill_formula");
      return true;
    },
  );

  const registrySource = await readFile(new URL("../src/tools/index.ts", import.meta.url), "utf8");
  assert.match(registrySource, /hostKind === "wps" \? \[createExecuteWpsJsTool\(\)\] : \[\]/u);
});
