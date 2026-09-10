import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@earendil-works/pi-agent-core";

import { createTraceDependenciesTool } from "../src/tools/trace-dependencies.ts";
import type { TraceDependenciesDetails } from "../src/tools/tool-details.ts";

interface CellState {
  value: unknown;
  formula: string;
}

class TraceRange {
  readonly numberFormat = [["General"]];
  readonly rowCount = 1;
  readonly columnCount = 1;
  readonly isNullObject = false;
  readonly format = {};
  private readonly cells: Map<string, CellState>;
  private readonly formulaGrid: unknown[][] | undefined;
  readonly address: string;

  constructor(cells: Map<string, CellState>, address: string, formulaGrid?: unknown[][]) {
    this.cells = cells;
    this.address = address;
    this.formulaGrid = formulaGrid;
  }

  load(_properties?: string): void {}

  get values(): unknown[][] {
    return [[this.cells.get(this.address)?.value ?? null]];
  }

  get formulas(): unknown[][] {
    return this.formulaGrid ?? [[this.cells.get(this.address)?.formula ?? ""]];
  }

  getDirectPrecedents(): never {
    throw new Error("Direct precedent API unavailable");
  }

  getDirectDependents(): never {
    throw new Error("Direct dependent API unavailable");
  }
}

class TraceSheet {
  readonly cells = new Map<string, CellState>();
  readonly name: string;
  private usedRange: { address: string; formulas: unknown[][] } | undefined;

  constructor(name: string) {
    this.name = name;
  }

  load(_properties?: string): void {}

  getRange(address: string): TraceRange {
    return new TraceRange(this.cells, address);
  }

  getUsedRangeOrNullObject(): TraceRange {
    return this.usedRange
      ? new TraceRange(this.cells, this.usedRange.address, this.usedRange.formulas)
      : new TraceRange(this.cells, "A1", [[""]]);
  }

  setUsedRange(address: string, formulas: unknown[][]): void {
    this.usedRange = { address, formulas };
  }
}

function createContext(sheets: TraceSheet[]): unknown {
  return {
    workbook: {
      worksheets: {
        items: sheets,
        load: (_properties?: string) => undefined,
        getActiveWorksheet: () => sheets[0],
        getItem: (name: string) => {
          const sheet = sheets.find((candidate) => candidate.name === name);
          if (!sheet) throw new Error(`Unknown sheet: ${name}`);
          return sheet;
        },
      },
    },
    sync: () => Promise.resolve(),
  };
}

async function withExcel<T>(context: unknown, action: () => Promise<T>): Promise<T> {
  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (host: unknown) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    return await action();
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

function details(result: AgentToolResult<TraceDependenciesDetails | undefined>): TraceDependenciesDetails {
  if (!result.details) throw new Error("Expected trace details.");
  return result.details;
}

void test("trace_dependencies defaults to precedents and parses formula references through the host", async () => {
  const calc = new TraceSheet("Calc");
  const input = new TraceSheet("Input Data");
  calc.cells.set("D10", {
    value: 42,
    formula: "=SUM(A1:B2,'Input Data'!$C$5,D1)+IF(A2=\"E9\",0,0)",
  });
  calc.cells.set("A1", { value: 10, formula: "" });
  calc.cells.set("D1", { value: 12, formula: "" });
  input.cells.set("C5", { value: 20, formula: "" });

  // The tool implementation returns this details shape but omits the second AgentTool generic.
  const result = await withExcel(createContext([calc, input]), () =>
    createTraceDependenciesTool().execute("trace-1", { cell: "Calc!D10", depth: 1 })) as
    AgentToolResult<TraceDependenciesDetails | undefined>;

  assert.deepEqual(details(result), {
    kind: "trace_dependencies",
    root: {
      address: "Calc!D10",
      value: 42,
      numberFormat: "General",
      formula: "=SUM(A1:B2,'Input Data'!$C$5,D1)+IF(A2=\"E9\",0,0)",
      precedents: [
        { address: "Calc!A1", value: 10, numberFormat: "General", precedents: [] },
        { address: "'Input Data'!C5", value: 20, numberFormat: "General", precedents: [] },
        { address: "Calc!D1", value: 12, numberFormat: "General", precedents: [] },
        { address: "Calc!A2", value: null, numberFormat: "General", precedents: [] },
      ],
    },
    mode: "precedents",
    maxDepth: 1,
    nodeCount: 5,
    edgeCount: 4,
    source: "formula_scan",
    truncated: false,
  });
});

void test("trace_dependencies finds dependents that reference the target inside a range", async () => {
  const calc = new TraceSheet("Calc");
  calc.cells.set("B2", { value: 8, formula: "" });
  calc.cells.set("C3", { value: 18, formula: "=SUM(A1:B2)" });
  calc.setUsedRange("A1:C3", [
    ["", "", ""],
    ["", "", ""],
    ["", "", "=SUM(A1:B2)"],
  ]);

  // The tool implementation returns this details shape but omits the second AgentTool generic.
  const result = await withExcel(createContext([calc]), () =>
    createTraceDependenciesTool().execute("trace-2", {
      cell: "Calc!B2",
      mode: "dependents",
      depth: 1,
    })) as AgentToolResult<TraceDependenciesDetails | undefined>;

  const trace = details(result);
  assert.equal(trace.mode, "dependents");
  assert.equal(trace.source, "formula_scan");
  assert.equal(trace.nodeCount, 2);
  assert.equal(trace.edgeCount, 1);
  assert.deepEqual(trace.root.precedents.map((node) => node.address), ["Calc!C3"]);
});
