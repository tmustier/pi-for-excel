import assert from "node:assert/strict";
import { test } from "node:test";

import { createExplainFormulaTool } from "../src/tools/explain-formula.ts";

function createRange(address: string, value: DynamicValue, formula: DynamicValue) {
  return {
    address,
    values: [[value]],
    formulas: [[formula]],
    load: (_properties: string): void => {},
  };
}

async function withFormulaWorkbook<T>(action: () => Promise<T>): Promise<T> {
  const ranges = new Map([
    ["D10", createRange("D10", 1520, "=IFERROR(SUM(A1:A3)+SUM(B1:B3)+C1,0)")],
    ["A1", createRange("A1:A3", "", "")],
    ["B1", createRange("B1:B3", "x".repeat(200), "")],
  ]);
  const sheet = {
    name: "Sales, Q1",
    load: (_properties: string): void => {},
    getRange: (address: string) => {
      const range = ranges.get(address);
      if (!range) throw new Error(`Unexpected range: ${address}`);
      return range;
    },
  };
  const context = {
    workbook: {
      worksheets: {
        getItem: (_name: string) => sheet,
        getActiveWorksheet: () => sheet,
      },
    },
    sync: (): Promise<void> => Promise.resolve(),
  };

  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (hostContext: DynamicValue) => Promise<TResult>): Promise<TResult> =>
      callback(context),
  });

  try {
    return await action();
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

void test("explain_formula explains a quoted-sheet formula and previews bounded references", async () => {
  await withFormulaWorkbook(async () => {
    const result = await createExplainFormulaTool().execute("call-1", {
      cell: "'Sales, Q1'!D10",
      max_references: 2,
    });

    assert.deepEqual(
      {
        ...result.details,
        references: result.details.references.map(({ address, valuePreview }) => ({
          address,
          valuePreview: valuePreview === "(blank)" ? valuePreview : "<long preview>",
        })),
      },
      {
        kind: "explain_formula",
        cell: "'Sales, Q1'!D10",
        hasFormula: true,
        formula: "=IFERROR(SUM(A1:A3)+SUM(B1:B3)+C1,0)",
        valuePreview: "1520",
        explanation:
          "Current value: 1520. The formula substitutes a fallback when an error occurs and adds values across 3 direct references. Reference preview is truncated; inspect cited cells for the complete lineage.",
        references: [
          { address: "'Sales, Q1'!A1:A3", valuePreview: "(blank)" },
          { address: "'Sales, Q1'!B1:B3", valuePreview: "<long preview>" },
        ],
        truncated: true,
      },
    );
    assert.match(result.details.references[1]?.valuePreview ?? "", /^x{119}…$/u);
  });
});

void test("explain_formula rejects ranges and multi-area addresses", async () => {
  const tool = createExplainFormulaTool();

  for (const cell of ["Sheet1!A1:B2", "A1,B2"]) {
    const result = await tool.execute("call-invalid", { cell });
    const text = result.content[0]?.type === "text" ? result.content[0].text : "";
    assert.equal(text, "Error: explain_formula expects a single cell, not a range.");
  }
});
