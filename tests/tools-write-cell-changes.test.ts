import assert from "node:assert/strict";
import { test } from "node:test";

import { createWriteCellsTool } from "../src/tools/write-cells.ts";

interface StoredCell {
  value: unknown;
  formula: string;
}

function columnName(index: number): string {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

class WriteRange {
  readonly address: string;
  readonly rowCount = 1;
  readonly columnCount: number;
  readonly numberFormat: unknown[][];
  private readonly cells: StoredCell[];

  constructor(cells: StoredCell[], address: string) {
    this.cells = cells;
    this.address = address;
    const [start = "A1", end = start] = address.split(":");
    const startMatch = /^([A-Z]+)(\d+)$/u.exec(start);
    const endMatch = /^([A-Z]+)(\d+)$/u.exec(end);
    if (!startMatch?.[1] || !endMatch?.[1]) throw new Error(`Unsupported test address: ${address}`);
    const toIndex = (letters: string): number => [...letters].reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
    this.columnCount = toIndex(endMatch[1]) - toIndex(startMatch[1]) + 1;
    this.numberFormat = [Array.from({ length: this.columnCount }, () => "General")];
  }

  load(_properties?: string | string[]): void {}

  get values(): unknown[][] {
    return [this.cells.slice(0, this.columnCount).map((cell) => cell.value)];
  }

  set values(values: unknown[][]) {
    for (let index = 0; index < this.columnCount; index += 1) {
      const value = values[0]?.[index] ?? null;
      const formula = typeof value === "string" && value.startsWith("=") ? value : "";
      this.cells[index] = { value: formula ? null : value, formula };
    }
  }

  get formulas(): unknown[][] {
    return [this.cells.slice(0, this.columnCount).map((cell) => cell.formula)];
  }
}

async function withWriteHost<T>(action: () => Promise<T>): Promise<T> {
  const cells = Array.from({ length: 13 }, (): StoredCell => ({ value: null, formula: "" }));
  const sheet = {
    name: "Sheet1",
    load: (_properties?: string | string[]): void => {},
    getRange: (address: string): WriteRange => new WriteRange(cells, address),
  };
  const context = {
    workbook: {
      worksheets: {
        getActiveWorksheet: () => sheet,
        getItem: (_name: string) => sheet,
      },
    },
    sync: (): Promise<void> => Promise.resolve(),
  };
  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (hostContext: unknown) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    return await action();
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

void test("write_cells reports value and formula changes with a bounded sample", async () => {
  await withWriteHost(async () => {
    const values = [["=2+2", 20, ...Array.from({ length: 11 }, (_, index) => index + 30)]];
    const result = await createWriteCellsTool().execute("write-1", {
      start_cell: "Sheet1!A1",
      values,
    });

    assert.equal(result.details.changes?.changedCount, 13);
    assert.equal(result.details.changes?.truncated, true);
    assert.equal(result.details.changes?.sample.length, 12);
    assert.deepEqual(result.details.changes?.sample.slice(0, 2), [
      {
        address: "Sheet1!A1",
        beforeValue: "",
        afterValue: "",
        afterFormula: "=2+2",
      },
      {
        address: "Sheet1!B1",
        beforeValue: "",
        afterValue: "20",
      },
    ]);
    assert.equal(result.details.changes?.sample.at(-1)?.address, `Sheet1!${columnName(11)}1`);
  });
});
