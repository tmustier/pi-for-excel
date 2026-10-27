import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { Type } from "typebox";

import {
  withWorkbookCoordinator,
  type WorkbookCoordinatorContextProvider,
  type WorkbookExecutionPolicy,
  type WorkbookMutationEvent,
} from "../src/tools/with-workbook-coordinator.ts";
import type {
  WorkbookCoordinator,
  WorkbookCoordinatorEvent,
  WorkbookOperationContext,
  WorkbookQueueSnapshot,
} from "../src/workbook/coordinator.ts";
import {
  createUnsupportedHostTool,
  UnsupportedHostToolError,
} from "../src/tools/unsupported-host-tool.ts";

const GENERIC_PARAMS = Type.Object({});

interface TestDetails {
  echo: string;
}

class FakeCoordinator implements WorkbookCoordinator {
  readonly readCalls: WorkbookOperationContext[] = [];
  readonly writeCalls: WorkbookOperationContext[] = [];

  private revision = 0;

  async runRead<T>(ctx: WorkbookOperationContext, fn: () => Promise<T>): Promise<T> {
    this.readCalls.push(ctx);
    return fn();
  }

  async runWrite<T>(ctx: WorkbookOperationContext, fn: () => Promise<T>): Promise<{ result: T; revision: number }> {
    this.writeCalls.push(ctx);
    const result = await fn();
    this.revision += 1;
    return { result, revision: this.revision };
  }

  getRevision(_workbookId: string): number {
    return this.revision;
  }

  getSnapshot(_workbookId: string): WorkbookQueueSnapshot {
    return {
      revision: this.revision,
      queuedWrites: 0,
      activeWrite: null,
      queued: [],
    };
  }

  subscribe(_listener: (event: WorkbookCoordinatorEvent) => void): () => void {
    return () => {};
  }
}

function createContextProvider(workbookId: string | null, sessionId = "session-1"): WorkbookCoordinatorContextProvider {
  return {
    getWorkbookId: () => Promise.resolve(workbookId),
    getSessionId: () => sessionId,
  };
}

function makeTool(
  name: string,
  executeImpl?: () => Promise<AgentToolResult<TestDetails>>,
): AgentTool<typeof GENERIC_PARAMS, TestDetails> {
  return {
    name,
    label: name,
    description: `test ${name}`,
    parameters: GENERIC_PARAMS,
    execute: async () => {
      if (executeImpl) {
        return executeImpl();
      }

      return {
        content: [{ type: "text", text: `${name}:ok` }],
        details: { echo: name },
      };
    },
  };
}

function wrapSingleTool(args: {
  tool: AgentTool<typeof GENERIC_PARAMS, TestDetails>;
  coordinator: WorkbookCoordinator;
  contextProvider: WorkbookCoordinatorContextProvider;
  mutationEvents: WorkbookMutationEvent[];
  invalidatedWorkbookIds: Array<string | null>;
  executionPolicy?: WorkbookExecutionPolicy;
}): AgentTool {
  const wrapped = withWorkbookCoordinator(
    [args.tool],
    args.coordinator,
    args.contextProvider,
    {
      onWriteCommitted: (event) => {
        args.mutationEvents.push(event);
        if (event.impact === "structure") {
          args.invalidatedWorkbookIds.push(event.workbookId);
        }
      },
    },
    args.executionPolicy,
  );

  const first = wrapped[0];
  if (!first) {
    throw new Error("Wrapped tool missing");
  }

  return first;
}

void test("modify_structure write emits structure-impact mutation event", async () => {
  const coordinator = new FakeCoordinator();
  const mutationEvents: WorkbookMutationEvent[] = [];
  const invalidatedWorkbookIds: Array<string | null> = [];

  const wrapped = wrapSingleTool({
    tool: makeTool("modify_structure"),
    coordinator,
    contextProvider: createContextProvider("url_sha256:abc"),
    mutationEvents,
    invalidatedWorkbookIds,
  });

  const result = await wrapped.execute("tc-1", { action: "add_sheet" });

  assert.equal(coordinator.readCalls.length, 0);
  assert.equal(coordinator.writeCalls.length, 1);
  assert.equal(mutationEvents.length, 1);
  assert.equal(mutationEvents[0]?.toolName, "modify_structure");
  assert.equal(mutationEvents[0]?.impact, "structure");
  assert.equal(mutationEvents[0]?.workbookId, "url_sha256:abc");
  assert.deepEqual(invalidatedWorkbookIds, ["url_sha256:abc"]);

  const firstBlock = result.content[0];
  assert.equal(firstBlock?.type, "text");
});

void test("unsupported-host tools bypass workbook coordinator and fail fast", async () => {
  const coordinator = new FakeCoordinator();
  const mutationEvents: WorkbookMutationEvent[] = [];
  const invalidatedWorkbookIds: Array<string | null> = [];
  let approvalCalls = 0;
  let executeCount = 0;

  const unsupportedTool = createUnsupportedHostTool(
    makeTool("format_cells", () => {
      executeCount += 1;
      return Promise.resolve({
        content: [{ type: "text", text: "should not execute" }],
        details: { echo: "format_cells" },
      });
    }),
    "wps",
  );

  const wrapped = wrapSingleTool({
    tool: unsupportedTool,
    coordinator,
    contextProvider: createContextProvider("wps_path_sha256:safe"),
    mutationEvents,
    invalidatedWorkbookIds,
    executionPolicy: {
      getExecutionMode: () => Promise.resolve("safe"),
      requestToolApproval: () => {
        approvalCalls += 1;
        return Promise.resolve(true);
      },
    },
  });

  assert.equal(wrapped, unsupportedTool);
  await assert.rejects(
    async () => wrapped.execute("tc-unsupported", { range: "Sheet1!A1", format: { bold: true } }),
    (error: unknown) => {
      assert.ok(error instanceof UnsupportedHostToolError);
      assert.equal(error.toolName, "format_cells");
      assert.equal(error.hostKind, "wps");
      return true;
    },
  );

  assert.equal(approvalCalls, 0);
  assert.equal(executeCount, 0);
  assert.equal(coordinator.readCalls.length, 0);
  assert.equal(coordinator.writeCalls.length, 0);
  assert.equal(mutationEvents.length, 0);
  assert.deepEqual(invalidatedWorkbookIds, []);
});

void test("safe execution mode blocks mutate calls when approval is denied", async () => {
  const coordinator = new FakeCoordinator();
  const mutationEvents: WorkbookMutationEvent[] = [];
  const invalidatedWorkbookIds: Array<string | null> = [];
  let approvalCalls = 0;

  const wrapped = wrapSingleTool({
    tool: makeTool("write_cells"),
    coordinator,
    contextProvider: createContextProvider("url_sha256:safe"),
    mutationEvents,
    invalidatedWorkbookIds,
    executionPolicy: {
      getExecutionMode: () => Promise.resolve("safe"),
      requestToolApproval: () => {
        approvalCalls += 1;
        return Promise.resolve(false);
      },
    },
  });

  await assert.rejects(
    () => wrapped.execute("tc-safe-block", { range: "Sheet1!A1", values: [[1]] }),
    /declined this workbook change/u,
  );

  assert.equal(approvalCalls, 1);
  assert.equal(coordinator.readCalls.length, 0);
  assert.equal(coordinator.writeCalls.length, 0);
  assert.equal(mutationEvents.length, 0);
  assert.deepEqual(invalidatedWorkbookIds, []);
});

void test("safe execution mode fails closed when no approval handler is configured", async () => {
  const coordinator = new FakeCoordinator();
  const mutationEvents: WorkbookMutationEvent[] = [];
  const invalidatedWorkbookIds: Array<string | null> = [];

  const wrapped = wrapSingleTool({
    tool: makeTool("write_cells"),
    coordinator,
    contextProvider: createContextProvider("url_sha256:safe-confirm"),
    mutationEvents,
    invalidatedWorkbookIds,
    executionPolicy: {
      getExecutionMode: () => Promise.resolve("safe"),
    },
  });

  await assert.rejects(
    () => wrapped.execute("tc-safe-confirm", { range: "Sheet1!A1", values: [[1]] }),
    /confirmation UI is unavailable/i,
  );

  assert.equal(coordinator.writeCalls.length, 0);
  assert.equal(mutationEvents.length, 0);
  assert.deepEqual(invalidatedWorkbookIds, []);
});

void test("safe execution mode skips approval prompts for read tools", async () => {
  const coordinator = new FakeCoordinator();
  const mutationEvents: WorkbookMutationEvent[] = [];
  const invalidatedWorkbookIds: Array<string | null> = [];
  let approvalCalls = 0;

  const wrapped = wrapSingleTool({
    tool: makeTool("read_range"),
    coordinator,
    contextProvider: createContextProvider("url_sha256:safe-read"),
    mutationEvents,
    invalidatedWorkbookIds,
    executionPolicy: {
      getExecutionMode: () => Promise.resolve("safe"),
      requestToolApproval: () => {
        approvalCalls += 1;
        return Promise.resolve(true);
      },
    },
  });

  await wrapped.execute("tc-safe-read", { range: "Sheet1!A1:A2" });

  assert.equal(approvalCalls, 0);
  assert.equal(coordinator.readCalls.length, 1);
  assert.equal(coordinator.writeCalls.length, 0);
  assert.equal(mutationEvents.length, 0);
  assert.deepEqual(invalidatedWorkbookIds, []);
});

void test("content-impact mutation tools do not trigger structure invalidation", async (t) => {
  const cases: Array<{ toolName: string; params: Record<string, unknown> }> = [
    { toolName: "write_cells", params: { range: "Sheet1!A1", values: [[1]] } },
    { toolName: "format_cells", params: { range: "Sheet1!A1", format: { bold: true } } },
    { toolName: "comments", params: { action: "delete", range: "Sheet1!A1" } },
    {
      toolName: "python_transform_range",
      params: { range: "Sheet1!A1:B3", code: "result = input_data['values']" },
    },
    { toolName: "view_settings", params: { action: "hide_gridlines" } },
    { toolName: "view_settings", params: { action: "activate", sheet: "Sheet1" } },
    { toolName: "workbook_history", params: { action: "restore", snapshot_id: "checkpoint-1" } },
  ];

  for (const item of cases) {
    const action = typeof item.params.action === "string" ? item.params.action : "";
    const testName = action ? `${item.toolName}:${action}` : item.toolName;

    await t.test(testName, async () => {
      const coordinator = new FakeCoordinator();
      const mutationEvents: WorkbookMutationEvent[] = [];
      const invalidatedWorkbookIds: Array<string | null> = [];

      const wrapped = wrapSingleTool({
        tool: makeTool(item.toolName),
        coordinator,
        contextProvider: createContextProvider("url_sha256:content"),
        mutationEvents,
        invalidatedWorkbookIds,
      });

      await wrapped.execute(`tc-${item.toolName}`, item.params);

      assert.equal(coordinator.readCalls.length, 0);
      assert.equal(coordinator.writeCalls.length, 1);
      assert.equal(mutationEvents.length, 1);
      assert.equal(mutationEvents[0]?.impact, "content");
      assert.deepEqual(invalidatedWorkbookIds, []);
    });
  }
});

void test("view_settings visibility actions trigger structure invalidation", async () => {
  const coordinator = new FakeCoordinator();
  const mutationEvents: WorkbookMutationEvent[] = [];
  const invalidatedWorkbookIds: Array<string | null> = [];

  const wrapped = wrapSingleTool({
    tool: makeTool("view_settings"),
    coordinator,
    contextProvider: createContextProvider("url_sha256:sheet-vis"),
    mutationEvents,
    invalidatedWorkbookIds,
  });

  await wrapped.execute("tc-view-settings-hide", { action: "hide_sheet", sheet: "Sheet1" });

  assert.equal(coordinator.readCalls.length, 0);
  assert.equal(coordinator.writeCalls.length, 1);
  assert.equal(mutationEvents.length, 1);
  assert.equal(mutationEvents[0]?.impact, "structure");
  assert.deepEqual(invalidatedWorkbookIds, ["url_sha256:sheet-vis"]);
});

void test("read-only tool paths never emit mutation events", async (t) => {
  const cases: Array<{ toolName: string; params: Record<string, unknown> }> = [
    { toolName: "read_range", params: { range: "Sheet1!A1:B2" } },
    { toolName: "comments", params: { action: "read", range: "Sheet1!A1" } },
    { toolName: "view_settings", params: { action: "get" } },
    { toolName: "workbook_history", params: { action: "list" } },
    { toolName: "tmux", params: { action: "list_sessions" } },
    { toolName: "python_run", params: { code: "print(1)" } },
    { toolName: "libreoffice_convert", params: { input_path: "/tmp/a.xlsx", target_format: "csv" } },
  ];

  for (const item of cases) {
    const action = typeof item.params.action === "string" ? item.params.action : "";
    const testName = action ? `${item.toolName}:${action}` : item.toolName;

    await t.test(testName, async () => {
      const coordinator = new FakeCoordinator();
      const mutationEvents: WorkbookMutationEvent[] = [];
      const invalidatedWorkbookIds: Array<string | null> = [];

      const wrapped = wrapSingleTool({
        tool: makeTool(item.toolName),
        coordinator,
        contextProvider: createContextProvider("url_sha256:read"),
        mutationEvents,
        invalidatedWorkbookIds,
      });

      await wrapped.execute(`tc-${item.toolName}`, item.params);

      assert.equal(coordinator.readCalls.length, 1);
      assert.equal(coordinator.writeCalls.length, 0);
      assert.equal(mutationEvents.length, 0);
      assert.deepEqual(invalidatedWorkbookIds, []);
    });
  }
});

void test("write failures do not emit mutation events", async () => {
  const coordinator = new FakeCoordinator();
  const mutationEvents: WorkbookMutationEvent[] = [];
  const invalidatedWorkbookIds: Array<string | null> = [];

  const wrapped = wrapSingleTool({
    tool: makeTool("modify_structure", () => Promise.reject(new Error("write failed"))),
    coordinator,
    contextProvider: createContextProvider("url_sha256:abc"),
    mutationEvents,
    invalidatedWorkbookIds,
  });

  await assert.rejects(() => wrapped.execute("tc-fail", { action: "add_sheet" }), /write failed/);

  assert.equal(coordinator.writeCalls.length, 1);
  assert.equal(mutationEvents.length, 0);
  assert.deepEqual(invalidatedWorkbookIds, []);
});


void test("a read issued after a write in the same batch sees the written value", async () => {
  let cell = "old";
  let finishWrite: () => void = () => {};
  const writeTool = makeTool("write_cells", () => new Promise((resolve) => {
    finishWrite = () => {
      cell = "new";
      resolve({ content: [{ type: "text", text: "written" }], details: { echo: "write" } });
    };
  }));
  const readTool = makeTool("read_range", () => Promise.resolve({
    content: [{ type: "text", text: cell }],
    details: { echo: cell },
  }));
  const [write, read, laterWrite] = withWorkbookCoordinator(
    [writeTool, readTool, makeTool("format_cells")],
    new FakeCoordinator(),
    createContextProvider("wb-1"),
  );
  if (!write || !read || !laterWrite) throw new Error("Wrapped tools missing");

  // The agent loop starts a batch's calls in order, then lets them run concurrently.
  const writing = write.execute("w", {});
  const reading = read.execute("r", {});
  await new Promise((resolve) => setTimeout(resolve, 10));
  finishWrite();

  const [, readResult] = await Promise.all([writing, reading]);
  assert.deepEqual(readResult.details, { echo: "new" });

  // A read issued before a write does not wait for it.
  let releaseLater: () => void = () => {};
  const blockingWrite = makeTool("write_cells", () => new Promise((resolve) => {
    releaseLater = () => resolve({ content: [{ type: "text", text: "late" }], details: { echo: "late" } });
  }));
  const [readFirst, writeSecond] = withWorkbookCoordinator(
    [readTool, blockingWrite],
    new FakeCoordinator(),
    createContextProvider("wb-1"),
  );
  if (!readFirst || !writeSecond) throw new Error("Wrapped tools missing");
  const earlyRead = readFirst.execute("r2", {});
  const pendingWrite = writeSecond.execute("w2", {});
  assert.deepEqual((await earlyRead).details, { echo: "new" });
  releaseLater();
  await pendingWrite;
});

void test("Stop tells the agent whether a mutation ran", async () => {
  const stopped = new AbortController();
  stopped.abort();
  const [beforeRun] = withWorkbookCoordinator([makeTool("format_cells")], new FakeCoordinator(), createContextProvider("wb-1"));
  if (!beforeRun) throw new Error("Wrapped tool missing");
  await assert.rejects(beforeRun.execute("s1", {}, stopped.signal), /Stopped before running; nothing changed in the workbook\./u);

  // Stop pressed while a one-batch mutation runs: its change stands, and the result says so.
  const during = new AbortController();
  const [whileRunning] = withWorkbookCoordinator(
    [makeTool("format_cells", () => {
      during.abort();
      return Promise.resolve({ content: [{ type: "text", text: "Formatted A1." }], details: { echo: "ok" } });
    })],
    new FakeCoordinator(),
    createContextProvider("wb-1"),
  );
  if (!whileRunning) throw new Error("Wrapped tool missing");
  const result = await whileRunning.execute("s2", {}, during.signal);
  const first = result.content[0];
  assert.equal(first?.type === "text" ? first.text : "", "Formatted A1.\n\n⏹ Stop was pressed while this ran; the change above was already applied and stands.");
});

void test("Stop settles a write Excel never answers, so the workbook lock is released", async (t) => {
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const coordinator = new FakeCoordinator();
  const windows: Array<{ toolCallId: string; settledAt: number }> = [];
  const [hung] = withWorkbookCoordinator(
    [makeTool("write_cells", () => new Promise(() => {}))],
    coordinator,
    createContextProvider("url_sha256:abc"),
    { onWriteStarted: (window) => windows.push(window) },
  );
  if (!hung) throw new Error("Wrapped tool missing");

  const controller = new AbortController();
  const call = hung.execute("tc-hung", {}, controller.signal);
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(windows[0]?.settledAt, Number.POSITIVE_INFINITY, "write window open while the call runs");
  controller.abort();
  t.mock.timers.tick(10_000);

  await assert.rejects(call, (error: unknown) => {
    assert.ok(error instanceof Error);
    assert.match(error.message, /Stopped by the user while Excel was still working.*outcome is unknown/u);
    assert.equal(Reflect.get(error, "outcomeUnknown"), true);
    return true;
  });
  assert.equal(windows[0]?.toolCallId, "tc-hung");
  assert.ok(Number.isFinite(windows[0]?.settledAt), "write window closed");
});

void test("a write that finishes a few seconds after Stop reports its own result", async (t) => {
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const [slow] = withWorkbookCoordinator(
    [makeTool("write_cells", () => new Promise((resolve) => {
      setTimeout(() => resolve({ content: [{ type: "text", text: "Stopped after writing Calc!K11:K20; backup fae5e5d3" }], details: { echo: "slow" } }), 6_000);
    }))],
    new FakeCoordinator(),
    createContextProvider("url_sha256:abc"),
  );
  if (!slow) throw new Error("Wrapped tool missing");

  const controller = new AbortController();
  const call = slow.execute("tc-slow", {}, controller.signal);
  await new Promise((resolve) => setImmediate(resolve));
  controller.abort();
  t.mock.timers.tick(6_000);

  const result = await call;
  const first = result.content[0];
  // #791's wrapper may append a note that the change stands; the tool's own report comes first.
  assert.ok((first?.type === "text" ? first.text : "").startsWith("Stopped after writing Calc!K11:K20; backup fae5e5d3"));
});

void test("a tool that reports its own Stop result is not cut short by the wrapper", async (t) => {
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const tool = Object.assign(makeTool("write_cells", () => new Promise((resolve) => {
    setTimeout(() => resolve({ content: [{ type: "text", text: "Stopped after writing rows 1-40 of 100; backup 5d2c0a1e" }], details: { echo: "slow" } }), 20_000);
  })), { reportsStop: true as const });
  const [wrapped] = withWorkbookCoordinator([tool], new FakeCoordinator(), createContextProvider("url_sha256:abc"));
  if (!wrapped) throw new Error("Wrapped tool missing");

  const controller = new AbortController();
  const call = wrapped.execute("tc-reports", {}, controller.signal);
  await new Promise((resolve) => setImmediate(resolve));
  controller.abort();
  t.mock.timers.tick(20_000);

  const first = (await call).content[0];
  assert.equal(first?.type === "text" ? first.text : "", "Stopped after writing rows 1-40 of 100; backup 5d2c0a1e");
});

void test("deleting backups asks even in Auto mode, before taking the workbook lock", async () => {
  const coordinator = new FakeCoordinator();
  const reasons: string[] = [];
  let ran = 0;

  const wrapped = wrapSingleTool({
    tool: makeTool("workbook_history", () => {
      ran += 1;
      return Promise.resolve({ content: [{ type: "text", text: "Cleared" }], details: { echo: "cleared" } });
    }),
    coordinator,
    contextProvider: createContextProvider("url_sha256:auto-history"),
    mutationEvents: [],
    invalidatedWorkbookIds: [],
    executionPolicy: {
      getExecutionMode: () => Promise.resolve("yolo"),
      requestToolApproval: (request) => {
        reasons.push(request.reason);
        // Nothing holds or waits for the lock while the user decides.
        assert.equal(coordinator.readCalls.length + coordinator.writeCalls.length, 0);
        return Promise.resolve(false);
      },
    },
  });

  await assert.rejects(() => wrapped.execute("tc-clear", { action: "clear" }), /declined/u);
  assert.deepEqual(reasons, ["always"]);
  assert.equal(ran, 0);

  // Listing backups never asks.
  await wrapped.execute("tc-list", { action: "list" });
  assert.deepEqual(reasons, ["always"]);
  assert.equal(ran, 1);
});

void test("Confirm mode does not ask about a write the tool would refuse anyway", async () => {
  const coordinator = new FakeCoordinator();
  let approvals = 0;
  let executed = 0;
  const refusing = {
    ...makeTool("write_cells", () => {
      executed += 1;
      return Promise.resolve({ content: [{ type: "text", text: "written" }], details: { echo: "written" } });
    }),
    checkBeforeApproval: (_id: string, params: { occupied?: boolean; allow?: boolean }) => Promise.resolve(
      params.occupied
        ? params.allow
          ? { overwrites: { count: 2, sample: [{ address: "Sheet1!A1", content: "=SUM(B1:B3)" }] } }
          : { refused: { content: [{ type: "text" as const, text: "⛔ Write blocked" }], details: { echo: "blocked" } } }
        : null,
    ),
  };
  const asked: unknown[] = [];
  const wrapped = wrapSingleTool({
    tool: refusing,
    coordinator,
    contextProvider: createContextProvider("url_sha256:precheck"),
    mutationEvents: [],
    invalidatedWorkbookIds: [],
    executionPolicy: {
      getExecutionMode: () => Promise.resolve("safe"),
      requestToolApproval: (request) => {
        approvals += 1;
        asked.push(request.overwrites);
        return Promise.resolve(true);
      },
    },
  });

  const blocked = await wrapped.execute("tc-blocked", { occupied: true });
  assert.match(JSON.stringify(blocked.content), /Write blocked/u);
  assert.equal(approvals, 0);
  assert.equal(executed, 0);
  assert.equal(coordinator.writeCalls.length, 0);

  await wrapped.execute("tc-ok", { occupied: false });
  assert.equal(approvals, 1);
  assert.equal(executed, 1);

  // An allowed overwrite asks, and the approval names what it replaces.
  await wrapped.execute("tc-overwrite", { occupied: true, allow: true });
  assert.deepEqual(asked, [undefined, { count: 2, sample: [{ address: "Sheet1!A1", content: "=SUM(B1:B3)" }] }]);
});
