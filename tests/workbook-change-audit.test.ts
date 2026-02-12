import assert from "node:assert/strict";
import { test } from "node:test";

import {
  buildChangeExplanation,
  MAX_EXPLANATION_PROMPT_CHARS,
  MAX_EXPLANATION_TEXT_CHARS,
} from "../src/audit/change-explanation.ts";
import {
  WorkbookChangeAuditLog,
  type AppendWorkbookChangeAuditEntryArgs,
} from "../src/audit/workbook-change-audit.ts";
import { createCommentsTool } from "../src/tools/comments.ts";
import { createViewSettingsTool } from "../src/tools/view-settings.ts";
import { createWorkbookHistoryTool } from "../src/tools/workbook-history.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import type { WorkbookRecoverySnapshot } from "../src/workbook/recovery-log.ts";

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
      if (value === undefined) {
        return Promise.resolve(null);
      }

      return Promise.resolve(value as T);
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function isUnknownArray(value: unknown): value is unknown[] {
  return Array.isArray(value);
}

function firstText(result: unknown): string {
  if (!isRecord(result)) {
    throw new Error("Expected tool result object");
  }

  const content = result.content;
  if (!isUnknownArray(content)) {
    throw new Error("Expected tool result content array");
  }

  const first = content[0];
  if (!isRecord(first) || first.type !== "text" || typeof first.text !== "string") {
    throw new Error("Expected first content block to be text");
  }

  return first.text;
}

function viewSettingsDetails(result: unknown): Record<string, unknown> {
  if (!isRecord(result)) {
    throw new Error("Expected tool result object");
  }

  const details = result.details;
  if (!isRecord(details)) {
    throw new Error("Expected tool result details object");
  }

  return details;
}

function createAuditCapture(): {
  entries: AppendWorkbookChangeAuditEntryArgs[];
  appendAuditEntry: (entry: AppendWorkbookChangeAuditEntryArgs) => Promise<void>;
} {
  const entries: AppendWorkbookChangeAuditEntryArgs[] = [];

  return {
    entries,
    appendAuditEntry: (entry: AppendWorkbookChangeAuditEntryArgs): Promise<void> => {
      entries.push(entry);
      return Promise.resolve();
    },
  };
}

function createRecoverySnapshot(id: string): WorkbookRecoverySnapshot {
  return {
    id,
    at: 1700000000300,
    toolName: "write_cells",
    toolCallId: "call-snapshot",
    address: "Sheet1!A1:A2",
    changedCount: 2,
    cellCount: 2,
    beforeValues: [[1], [2]],
    beforeFormulas: [[""], [""]],
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

void test("workbook change audit log accepts non-cell mutation tool entries", async () => {
  const settingsStore = createInMemorySettingsStore();

  const log = new WorkbookChangeAuditLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: "url_sha256:test789",
      workbookName: "Ops.xlsx",
      source: "document.url",
    }),
    now: () => 1700000000100,
    createId: () => "entry-ops",
  });

  await log.append({
    toolName: "modify_structure",
    toolCallId: "call-3",
    blocked: false,
    changedCount: 2,
    changes: [],
    summary: "inserted 2 row(s)",
  });

  const entries = await log.list();
  assert.equal(entries.length, 1);
  assert.equal(entries[0]?.toolName, "modify_structure");
  assert.equal(entries[0]?.summary, "inserted 2 row(s)");
  assert.equal(entries[0]?.changedCount, 2);
});

void test("workbook change audit log accepts comments/view_settings/workbook_history entries", async () => {
  const settingsStore = createInMemorySettingsStore();

  const log = new WorkbookChangeAuditLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: "url_sha256:test900",
      workbookName: "Controls.xlsx",
      source: "document.url",
    }),
  });

  await log.append({
    toolName: "comments",
    toolCallId: "call-comments",
    blocked: false,
    outputAddress: "Sheet1!A1",
    changedCount: 1,
    changes: [],
    summary: "updated comment at Sheet1!A1",
  });

  await log.append({
    toolName: "view_settings",
    toolCallId: "call-view",
    blocked: false,
    outputAddress: "Sheet1",
    changedCount: 1,
    changes: [],
    summary: "hid gridlines on Sheet1",
  });

  await log.append({
    toolName: "workbook_history",
    toolCallId: "call-history",
    blocked: true,
    changedCount: 0,
    changes: [],
    summary: "error: no recovery checkpoints available to restore",
  });

  const entries = await log.list();
  assert.equal(entries.length, 3);
  assert.deepEqual(entries.map((entry) => entry.toolName), ["workbook_history", "view_settings", "comments"]);
});

void test("comments tool appends audit entries for mutating actions", async () => {
  const auditCapture = createAuditCapture();

  const tool = createCommentsTool({
    dispatchAction: () => Promise.resolve({
      text: "Updated comment on **Sheet1!A1**: \"Approved\"",
      outputAddress: "Sheet1!A1",
      changedCount: 1,
      summary: "updated comment at Sheet1!A1",
    }),
    appendAuditEntry: auditCapture.appendAuditEntry,
    captureCommentThread: () => Promise.resolve({
      exists: true,
      content: "Before",
      resolved: false,
      replies: [],
    }),
    appendRecoverySnapshot: () => Promise.resolve({
      id: "checkpoint-1",
      at: 1700000000000,
      toolName: "comments",
      toolCallId: "tool-call-comments-success",
      address: "Sheet1!A1",
      changedCount: 1,
      cellCount: 1,
      beforeValues: [],
      beforeFormulas: [],
      snapshotKind: "comment_thread",
      commentThreadState: {
        exists: true,
        content: "Before",
        resolved: false,
        replies: [],
      },
    }),
  });

  const result = await tool.execute("tool-call-comments-success", {
    action: "update",
    range: "Sheet1!A1",
    content: "Approved",
  });

  assert.match(firstText(result), /Updated comment/u);
  if (isRecord(result.details)) {
    assert.equal(result.details.kind, "comments");
    if (isRecord(result.details.recovery)) {
      assert.equal(result.details.recovery.status, "checkpoint_created");
    } else {
      throw new Error("Expected recovery metadata for comments mutation");
    }
  } else {
    throw new Error("Expected comments details in tool result");
  }

  assert.equal(auditCapture.entries.length, 1);
  assert.equal(auditCapture.entries[0]?.toolName, "comments");
  assert.equal(auditCapture.entries[0]?.blocked, false);
  assert.equal(auditCapture.entries[0]?.outputAddress, "Sheet1!A1");
  assert.equal(auditCapture.entries[0]?.changedCount, 1);
});

void test("comments tool appends blocked audit entry on mutate validation error", async () => {
  const auditCapture = createAuditCapture();

  const tool = createCommentsTool({
    appendAuditEntry: auditCapture.appendAuditEntry,
  });

  const result = await tool.execute("tool-call-comments-error", {
    action: "add",
    range: "Sheet1!A1",
  });

  assert.match(firstText(result), /content is required for add/u);
  assert.equal(auditCapture.entries.length, 1);
  assert.equal(auditCapture.entries[0]?.toolName, "comments");
  assert.equal(auditCapture.entries[0]?.blocked, true);
  assert.equal(auditCapture.entries[0]?.changedCount, 0);
  assert.equal(auditCapture.entries[0]?.outputAddress, "Sheet1!A1");
});

void test("view_settings tool appends audit entries for mutate success and failure", async () => {
  const successAuditCapture = createAuditCapture();
  const successTool = createViewSettingsTool({
    executeAction: () => Promise.resolve({
      text: "Activated sheet \"Sheet2\".",
      outputAddress: "Sheet2",
      changedCount: 1,
      summary: "activated sheet Sheet2",
    }),
    appendAuditEntry: successAuditCapture.appendAuditEntry,
  });

  const successResult = await successTool.execute("tool-call-view-success", {
    action: "activate",
    sheet: "Sheet2",
  });

  assert.match(firstText(successResult), /Activated sheet/u);
  assert.match(firstText(successResult), /Recovery checkpoint not created/u);
  assert.equal(successAuditCapture.entries.length, 1);
  assert.equal(successAuditCapture.entries[0]?.toolName, "view_settings");
  assert.equal(successAuditCapture.entries[0]?.blocked, false);
  assert.equal(successAuditCapture.entries[0]?.outputAddress, "Sheet2");

  const successDetails = viewSettingsDetails(successResult);
  assert.equal(successDetails.kind, "view_settings");
  assert.equal(successDetails.action, "activate");
  assert.equal(successDetails.address, "Sheet2");
  assert.equal(
    isRecord(successDetails.recovery) ? successDetails.recovery.status : undefined,
    "not_available",
  );

  const errorAuditCapture = createAuditCapture();
  const errorTool = createViewSettingsTool({
    executeAction: () => Promise.reject(new Error("sheet is required for activate")),
    appendAuditEntry: errorAuditCapture.appendAuditEntry,
  });

  const errorResult = await errorTool.execute("tool-call-view-error", {
    action: "activate",
  });

  assert.match(firstText(errorResult), /sheet is required for activate/u);
  assert.match(firstText(errorResult), /Recovery checkpoint not created/u);
  assert.equal(errorAuditCapture.entries.length, 1);
  assert.equal(errorAuditCapture.entries[0]?.toolName, "view_settings");
  assert.equal(errorAuditCapture.entries[0]?.blocked, true);
  assert.equal(errorAuditCapture.entries[0]?.changedCount, 0);

  const errorDetails = viewSettingsDetails(errorResult);
  assert.equal(errorDetails.kind, "view_settings");
  assert.equal(errorDetails.action, "activate");
  assert.equal(errorDetails.address, undefined);
  assert.equal(
    isRecord(errorDetails.recovery) ? errorDetails.recovery.status : undefined,
    "not_available",
  );
});

void test("workbook_history restore appends audit entries for success and missing snapshot", async () => {
  const successAuditCapture = createAuditCapture();
  const successTool = createWorkbookHistoryTool({
    getRecoveryLog: () => ({
      listForCurrentWorkbook: () => Promise.resolve([createRecoverySnapshot("snapshot-1")]),
      restore: (snapshotId: string) => Promise.resolve({
        restoredSnapshotId: snapshotId,
        inverseSnapshotId: "inverse-1",
        address: "Sheet1!A1:A2",
        changedCount: 2,
      }),
      delete: () => Promise.resolve(false),
      clearForCurrentWorkbook: () => Promise.resolve(0),
    }),
    appendAuditEntry: successAuditCapture.appendAuditEntry,
  });

  const successResult = await successTool.execute("tool-call-history-success", {
    action: "restore",
  });

  assert.match(firstText(successResult), /Restored checkpoint/u);
  assert.equal(successAuditCapture.entries.length, 1);
  assert.equal(successAuditCapture.entries[0]?.toolName, "workbook_history");
  assert.equal(successAuditCapture.entries[0]?.blocked, false);
  assert.equal(successAuditCapture.entries[0]?.outputAddress, "Sheet1!A1:A2");
  assert.equal(successAuditCapture.entries[0]?.changedCount, 2);

  const missingAuditCapture = createAuditCapture();
  const missingTool = createWorkbookHistoryTool({
    getRecoveryLog: () => ({
      listForCurrentWorkbook: () => Promise.resolve([]),
      restore: () => Promise.resolve({
        restoredSnapshotId: "",
        inverseSnapshotId: null,
        address: "",
        changedCount: 0,
      }),
      delete: () => Promise.resolve(false),
      clearForCurrentWorkbook: () => Promise.resolve(0),
    }),
    appendAuditEntry: missingAuditCapture.appendAuditEntry,
  });

  const missingResult = await missingTool.execute("tool-call-history-missing", {
    action: "restore",
  });

  assert.match(firstText(missingResult), /No recovery checkpoints available to restore/u);
  assert.equal(missingAuditCapture.entries.length, 1);
  assert.equal(missingAuditCapture.entries[0]?.toolName, "workbook_history");
  assert.equal(missingAuditCapture.entries[0]?.blocked, true);
  assert.equal(missingAuditCapture.entries[0]?.changedCount, 0);
});

void test("buildChangeExplanation shapes citations and prompt from change metadata", () => {
  const explanation = buildChangeExplanation({
    toolName: "write_cells",
    blocked: false,
    changedCount: 4,
    summary: "wrote budget deltas",
    outputAddress: "Sheet1!A1:B2",
    changes: {
      changedCount: 4,
      truncated: false,
      sample: [
        {
          address: "Sheet1!A1",
          beforeValue: "10",
          afterValue: "12",
        },
        {
          address: "Sheet1!B2",
          beforeValue: "=SUM(B3:B10)",
          afterValue: "=SUM(B3:B11)",
          beforeFormula: "=SUM(B3:B10)",
          afterFormula: "=SUM(B3:B11)",
        },
      ],
    },
  });

  assert.match(explanation.prompt, /Tool: write_cells/u);
  assert.match(explanation.prompt, /Sample changes:/u);
  assert.match(explanation.prompt, /Sheet1!A1/u);

  assert.match(explanation.text, /changed 4 cell/u);
  assert.match(explanation.text, /Inspect:/u);
  assert.ok(explanation.citations.includes("Sheet1!A1"));
  assert.ok(explanation.citations.includes("Sheet1!B2"));
  assert.ok(explanation.citations.includes("Sheet1!A1:B2"));
  assert.equal(explanation.usedFallback, false);
});

void test("buildChangeExplanation enforces prompt and text budgets", () => {
  const longSummary = "x".repeat(MAX_EXPLANATION_PROMPT_CHARS + 200);

  const explanation = buildChangeExplanation({
    toolName: "python_transform_range",
    blocked: false,
    changedCount: 128,
    summary: longSummary,
    outputAddress: "Sheet1!A1:Z99",
    changes: {
      changedCount: 128,
      truncated: true,
      sample: [...Array(10).keys()].map((index) => ({
        address: `Sheet1!A${index + 1}`,
        beforeValue: "before-" + "v".repeat(120),
        afterValue: "after-" + "w".repeat(120),
      })),
    },
  });

  assert.equal(explanation.prompt.length <= MAX_EXPLANATION_PROMPT_CHARS, true);
  assert.equal(explanation.text.length <= MAX_EXPLANATION_TEXT_CHARS, true);
  assert.equal(explanation.truncated, true);
});

void test("buildChangeExplanation falls back gracefully on sparse metadata", () => {
  const explanation = buildChangeExplanation({
    toolName: "modify_structure",
    blocked: false,
  });

  assert.equal(explanation.citations.length, 0);
  assert.equal(explanation.usedFallback, true);
  assert.match(explanation.text, /Not enough audit metadata/u);
});
