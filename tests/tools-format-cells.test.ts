import assert from "node:assert/strict";
import { test } from "node:test";

import { initAppStorage } from "../src/storage/init-app-storage.ts";
import type { StorageBackend, StorageTransaction } from "../src/storage/local/types.ts";
import { createFormatCellsTool } from "../src/tools/format-cells.ts";

class MemoryStorageBackend implements StorageBackend {
  private readonly stores = new Map<string, Map<string, DynamicValue>>();

  private store(name: string): Map<string, DynamicValue> {
    const existing = this.stores.get(name);
    if (existing) return existing;
    const created = new Map<string, DynamicValue>();
    this.stores.set(name, created);
    return created;
  }

  get<T = DynamicValue>(storeName: string, key: string): Promise<T | null> {
    const value = this.store(storeName).get(key);
    // This boundary mirrors IndexedDB: callers own the requested persisted type.
    return Promise.resolve(value === undefined ? null : structuredClone(value) as T);
  }

  set<T = DynamicValue>(storeName: string, key: string, value: T): Promise<void> {
    this.store(storeName).set(key, structuredClone(value));
    return Promise.resolve();
  }

  delete(storeName: string, key: string): Promise<void> {
    this.store(storeName).delete(key);
    return Promise.resolve();
  }

  keys(storeName: string, prefix?: string): Promise<string[]> {
    const keys = [...this.store(storeName).keys()];
    return Promise.resolve(prefix ? keys.filter((key) => key.startsWith(prefix)) : keys);
  }

  getAllFromIndex<T = DynamicValue>(): Promise<T[]> {
    return Promise.resolve([]);
  }

  clear(storeName: string): Promise<void> {
    this.store(storeName).clear();
    return Promise.resolve();
  }

  has(storeName: string, key: string): Promise<boolean> {
    return Promise.resolve(this.store(storeName).has(key));
  }

  transaction<T>(
    _storeNames: string[],
    _mode: "readonly" | "readwrite",
    operation: (transaction: StorageTransaction) => Promise<T>,
  ): Promise<T> {
    return operation({
      get: <V = DynamicValue>(storeName: string, key: string) => this.get<V>(storeName, key),
      set: <V = DynamicValue>(storeName: string, key: string, value: V) => this.set(storeName, key, value),
      delete: (storeName: string, key: string) => this.delete(storeName, key),
    });
  }

  getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> {
    return Promise.resolve({ usage: 0, quota: 0, percent: 0 });
  }

  requestPersistence(): Promise<boolean> {
    return Promise.resolve(true);
  }
}

class BorderState {
  style = "Continuous";
  weight = "Thin";
  color = "#000000";
  load(_properties?: string): void {}
}

class FormatRange {
  readonly address = "A1";
  readonly rowCount = 1;
  readonly columnCount = 1;
  numberFormat = [["General"]];
  readonly borders = new Map<string, BorderState>();
  readonly format = {
    font: {
      color: "#000000",
      bold: false,
      italic: false,
      underline: "None",
      name: "Arial",
      size: 10,
      load: (_properties?: string) => undefined,
    },
    fill: { color: "#FFFFFF", load: (_properties?: string) => undefined },
    borders: {
      getItem: (edge: string): BorderState => {
        const existing = this.borders.get(edge);
        if (existing) return existing;
        const created = new BorderState();
        this.borders.set(edge, created);
        return created;
      },
    },
    load: (_properties?: string) => undefined,
  };

  load(_properties?: string): void {}
}

async function withFormatHost<T>(range: FormatRange, action: () => Promise<T>): Promise<T> {
  initAppStorage("tools-format-cells", new MemoryStorageBackend());
  const sheet = {
    name: "Sheet1",
    load: (_properties?: string) => undefined,
    getRange: (_address: string) => range,
  };
  const context = {
    workbook: {
      worksheets: {
        getActiveWorksheet: () => sheet,
        getItem: (_name: string) => sheet,
      },
    },
    sync: () => Promise.resolve(),
  };
  const hadExcel = Reflect.has(globalThis, "Excel");
  const previousExcel = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <TResult>(callback: (host: DynamicValue) => Promise<TResult>): Promise<TResult> => callback(context),
  });

  try {
    return await action();
  } finally {
    if (hadExcel) Reflect.set(globalThis, "Excel", previousExcel);
    else Reflect.deleteProperty(globalThis, "Excel");
  }
}

function firstText(result: { content: Array<{ type: string; text: string }> }): string {
  const block = result.content[0];
  if (!block || block.type !== "text") throw new Error("Expected text result.");
  return block.text;
}

void test("format_cells clears every border edge for the none shorthand", async () => {
  const range = new FormatRange();

  const result = await withFormatHost(range, () =>
    createFormatCellsTool().execute("format-1", { range: "A1", borders: "none" }));

  assert.match(firstText(result), /Formatted \*\*Sheet1!A1\*\*: none borders/u);
  assert.deepEqual(
    [...range.borders.entries()].map(([edge, border]) => [edge, border.style]),
    [
      ["EdgeTop", "None"],
      ["EdgeBottom", "None"],
      ["EdgeLeft", "None"],
      ["EdgeRight", "None"],
      ["InsideHorizontal", "None"],
      ["InsideVertical", "None"],
    ],
  );
});

void test("format_cells rejects an invalid border value before changing the host", async () => {
  const range = new FormatRange();
  const params = { range: "A1", borders: "remove-all" };

  // Runtime tool input originates from model JSON, so this deliberately bypasses compile-time schema checking.
  const result = await withFormatHost(range, () => createFormatCellsTool().execute("format-2", params));

  assert.match(firstText(result), /Error formatting: Invalid borders/u);
  assert.deepEqual([...range.borders], []);
});
