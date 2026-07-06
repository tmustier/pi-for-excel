import assert from "node:assert/strict";
import { test } from "node:test";

import { WorkbookRecoveryLog } from "../src/workbook/recovery-log.ts";
import type { WorkbookContext } from "../src/workbook/context.ts";
import { type RecoveryCommentThreadState } from "../src/workbook/recovery-states.ts";
import {
  RECOVERY_SETTING_KEY,
  createInMemorySettingsStore,
  findSnapshotById,
} from "./recovery-log-test-helpers.test.ts";

void test("restore rejects legacy snapshots without workbook identity", async () => {
  const settingsStore = createInMemorySettingsStore();

  await settingsStore.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [
      {
        id: "legacy-1",
        at: 1700000000000,
        toolName: "write_cells",
        toolCallId: "call-legacy",
        address: "Sheet1!A1",
        changedCount: 1,
        cellCount: 1,
        beforeValues: [["before"]],
        beforeFormulas: [["before"]],
      },
    ],
  });

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: "url_sha256:target",
      workbookName: "Target.xlsx",
      source: "document.url",
    }),
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  await assert.rejects(
    () => log.restore("legacy-1"),
    /missing workbook identity/i,
  );
});
void test("restore rejects when current workbook identity is unavailable", async () => {
  const settingsStore = createInMemorySettingsStore();

  await settingsStore.set(RECOVERY_SETTING_KEY, {
    version: 1,
    snapshots: [
      {
        id: "snap-1",
        at: 1700000000000,
        toolName: "write_cells",
        toolCallId: "call-1",
        address: "Sheet1!A1",
        changedCount: 1,
        cellCount: 1,
        beforeValues: [["before"]],
        beforeFormulas: [["before"]],
        workbookId: "url_sha256:origin",
      },
    ],
  });

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: (): Promise<WorkbookContext> => Promise.resolve({
      workbookId: null,
      workbookName: null,
      source: "unknown",
    }),
    applySnapshot: () => Promise.resolve({ values: [["old"]], formulas: [["old"]] }),
  });

  await assert.rejects(
    () => log.restore("snap-1"),
    /identity is unavailable/i,
  );
});
void test("restore applies checkpoint values and creates inverse checkpoint", async () => {
  const settingsStore = createInMemorySettingsStore();

  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:workbook-b",
    workbookName: "Model.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-${idCounter}`;
  };

  let appliedAddress = "";
  let appliedValues: DynamicValue[][] = [];

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000001000,
    createId,
    applySnapshot: (address: string, values: DynamicValue[][]) => {
      appliedAddress = address;
      appliedValues = values;
      return Promise.resolve({
        values: [[42]],
        formulas: [[42]],
      });
    },
  });

  const appended = await log.append({
    toolName: "fill_formula",
    toolCallId: "call-2",
    address: "Sheet2!B4",
    changedCount: 1,
    beforeValues: [[10]],
    beforeFormulas: [["=A1+A2"]],
  });

  assert.ok(appended);

  const restored = await log.restore(appended?.id ?? "");

  assert.equal(restored.address, "Sheet2!B4");
  assert.equal(restored.restoredSnapshotId, appended?.id);
  assert.equal(appliedAddress, "Sheet2!B4");
  assert.deepEqual(appliedValues, [["=A1+A2"]]);

  const snapshots = await log.listForCurrentWorkbook(10);
  assert.equal(snapshots.length, 2);

  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse?.toolName, "restore_snapshot");
  assert.equal(inverse?.restoredFromSnapshotId, appended?.id);
});
void test("restore applies comment-thread checkpoints and creates inverse checkpoint", async () => {
  const settingsStore = createInMemorySettingsStore();

  const workbookContext: WorkbookContext = {
    workbookId: "url_sha256:workbook-comments",
    workbookName: "Comments.xlsx",
    source: "document.url",
  };

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-comment-${idCounter}`;
  };

  let appliedAddress = "";
  let appliedState: DynamicValue = null;

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext: () => Promise.resolve(workbookContext),
    now: () => 1700000003000,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
    applyCommentThreadSnapshot: (address, state) => {
      appliedAddress = address;
      appliedState = state;
      return Promise.resolve({
        exists: true,
        content: "Current comment",
        resolved: false,
        replies: ["Current reply"],
      });
    },
  });

  const appended = await log.appendCommentThread({
    toolName: "comments",
    toolCallId: "call-comment",
    address: "Sheet1!C3",
    changedCount: 1,
    commentThreadState: {
      exists: true,
      content: "Original comment",
      resolved: true,
      replies: ["Original reply"],
    },
  });

  assert.ok(appended);

  const restored = await log.restore(appended?.id ?? "");

  assert.equal(restored.address, "Sheet1!C3");
  assert.equal(restored.restoredSnapshotId, appended?.id);
  assert.equal(appliedAddress, "Sheet1!C3");
  assert.deepEqual(appliedState, {
    exists: true,
    content: "Original comment",
    resolved: true,
    replies: ["Original reply"],
  });

  const snapshots = await log.listForCurrentWorkbook(10);
  const inverse = restored.inverseSnapshotId
    ? findSnapshotById(snapshots, restored.inverseSnapshotId)
    : null;

  assert.ok(inverse);
  assert.equal(inverse?.toolName, "restore_snapshot");
  assert.equal(inverse?.snapshotKind, "comment_thread");
  assert.equal(inverse?.restoredFromSnapshotId, appended?.id);
});
void test("restore round-trips comment-thread states for present and absent threads", async () => {
  const scenarios: ReadonlyArray<{
    name: string;
    targetState: RecoveryCommentThreadState;
    currentState: RecoveryCommentThreadState;
  }> = [
    {
      name: "present",
      targetState: {
        exists: true,
        content: "Original thread",
        resolved: false,
        replies: ["Reply A", "Reply B"],
      },
      currentState: {
        exists: true,
        content: "Current thread",
        resolved: true,
        replies: ["Current reply"],
      },
    },
    {
      name: "absent",
      targetState: {
        exists: false,
        content: "",
        resolved: false,
        replies: [],
      },
      currentState: {
        exists: true,
        content: "Current thread before delete",
        resolved: false,
        replies: ["Keep me"],
      },
    },
  ];

  for (let index = 0; index < scenarios.length; index += 1) {
    const scenario = scenarios[index];
    if (!scenario) {
      throw new Error("Expected comment scenario.");
    }

    const settingsStore = createInMemorySettingsStore();

    const workbookContext: WorkbookContext = {
      workbookId: `url_sha256:workbook-comment-roundtrip-${index}`,
      workbookName: "CommentsRoundtrip.xlsx",
      source: "document.url",
    };

    let idCounter = 0;
    const createId = (): string => {
      idCounter += 1;
      return `snap-comment-roundtrip-${index}-${idCounter}`;
    };

    let appliedAddress = "";
    let appliedState: RecoveryCommentThreadState | null = null;

    const log = new WorkbookRecoveryLog({
      getSettingsStore: () => Promise.resolve(settingsStore),
      getWorkbookContext: () => Promise.resolve(workbookContext),
      now: () => 1700000003200 + index,
      createId,
      applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
      applyCommentThreadSnapshot: (address, state) => {
        appliedAddress = address;
        appliedState = state;
        return Promise.resolve(scenario.currentState);
      },
    });

    const address = `Sheet1!C${index + 10}`;
    const appended = await log.appendCommentThread({
      toolName: "comments",
      toolCallId: `call-comment-roundtrip-${scenario.name}`,
      address,
      changedCount: 1,
      commentThreadState: scenario.targetState,
    });

    assert.ok(appended, `Expected appended comment checkpoint for ${scenario.name}.`);
    if (!appended) {
      throw new Error(`Expected appended comment checkpoint for ${scenario.name}.`);
    }

    const restored = await log.restore(appended.id);

    assert.equal(restored.address, address, `Expected restored address for ${scenario.name}.`);
    assert.equal(
      restored.restoredSnapshotId,
      appended.id,
      `Expected restored snapshot id for ${scenario.name}.`,
    );
    assert.equal(appliedAddress, address, `Expected apply address for ${scenario.name}.`);
    assert.deepEqual(appliedState, scenario.targetState, `Expected applied state for ${scenario.name}.`);

    const snapshots = await log.listForCurrentWorkbook(10);
    const inverse = restored.inverseSnapshotId
      ? findSnapshotById(snapshots, restored.inverseSnapshotId)
      : null;

    assert.ok(inverse, `Expected inverse snapshot for ${scenario.name}.`);
    assert.equal(inverse?.snapshotKind, "comment_thread", `Expected comment kind for ${scenario.name}.`);
    assert.equal(inverse?.restoredFromSnapshotId, appended.id, `Expected inverse source for ${scenario.name}.`);
    assert.deepEqual(
      inverse?.commentThreadState,
      scenario.currentState,
      `Expected inverse state for ${scenario.name}.`,
    );
  }
});
void test("restore rejects checkpoints from another workbook", async () => {
  const settingsStore = createInMemorySettingsStore();

  let currentWorkbookId: string | null = "url_sha256:workbook-src";
  const getWorkbookContext = (): Promise<WorkbookContext> => Promise.resolve({
    workbookId: currentWorkbookId,
    workbookName: "Workbook",
    source: currentWorkbookId ? "document.url" : "unknown",
  });

  let idCounter = 0;
  const createId = (): string => {
    idCounter += 1;
    return `snap-${idCounter}`;
  };

  const log = new WorkbookRecoveryLog({
    getSettingsStore: () => Promise.resolve(settingsStore),
    getWorkbookContext,
    createId,
    applySnapshot: () => Promise.resolve({ values: [[1]], formulas: [[1]] }),
  });

  const snapshot = await log.append({
    toolName: "write_cells",
    toolCallId: "call-5",
    address: "Sheet1!B2",
    beforeValues: [["before"]],
    beforeFormulas: [["before"]],
  });

  assert.ok(snapshot);

  currentWorkbookId = "url_sha256:workbook-other";

  await assert.rejects(
    async () => log.restore(snapshot?.id ?? ""),
    /different workbook/i,
  );
});
