import assert from "node:assert/strict";
import { test } from "node:test";

import { WorkbookChangeAuditLog } from "../src/audit/workbook-change-audit.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";

interface InMemorySettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
  delete(key: string): Promise<void>;
}

function createInMemorySettingsStore(): InMemorySettingsStore {
  const values = new Map<string, unknown>();

  return {
    get: <T>(key: string): Promise<T | null> => {
      const value = values.get(key);
      return Promise.resolve(value === undefined ? null : value as T);
    },
    set: (key: string, value: unknown): Promise<void> => {
      values.set(key, value);
      return Promise.resolve();
    },
    delete: (key: string): Promise<void> => {
      values.delete(key);
      return Promise.resolve();
    },
  };
}

void test("workbook change audit log appends and reloads entries", async () => {
  const settingsStore = createInMemorySettingsStore();

  const getWorkbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: "url_sha256:test123",
    workbookName: "Budget.xlsx",
    source: "document.url",
  });

  const logA = new WorkbookChangeAuditLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    now: () => 1700000000000,
    createId: () => "entry-1",
  });

  await logA.append({
    toolName: "write_cells",
    toolCallId: "call-1",
    blocked: false,
    outputAddress: "Sheet1!A1:B2",
    changedCount: 1,
    changes: [{
      address: "Sheet1!A1",
      beforeValue: "1",
      afterValue: "2",
    }],
  });

  const entriesA = await logA.list();
  assert.equal(entriesA.length, 1);
  assert.equal(entriesA[0]?.toolName, "write_cells");
  assert.equal(entriesA[0]?.workbookId, "url_sha256:test123");
  assert.equal(entriesA[0]?.workbookLabel, "Budget.xlsx");

  const logB = new WorkbookChangeAuditLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
  });

  const entriesB = await logB.list();
  assert.equal(entriesB.length, 1);
  assert.equal(entriesB[0]?.toolCallId, "call-1");
});

void test("workbook change audit log clear removes persisted entries", async () => {
  const settingsStore = createInMemorySettingsStore();

  const log = new WorkbookChangeAuditLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: "url_sha256:test456",
      workbookName: "Model.xlsx",
      source: "document.url",
    }),
  });

  await log.append({
    toolName: "fill_formula",
    toolCallId: "call-2",
    blocked: true,
    outputAddress: "Sheet1!C1:C5",
    changedCount: 0,
    changes: [],
  });

  assert.equal((await log.list()).length, 1);

  await log.clear();
  assert.equal((await log.list()).length, 0);
});
