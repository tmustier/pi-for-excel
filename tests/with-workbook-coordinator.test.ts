import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type } from "@sinclair/typebox";

import {
  withWorkbookCoordinator,
  type WorkbookCoordinatorContextProvider,
  type WorkbookMutationEvent,
} from "../src/tools/with-workbook-coordinator.ts";
import type {
  WorkbookCoordinator,
  WorkbookCoordinatorEvent,
  WorkbookOperationContext,
  WorkbookQueueSnapshot,
} from "../src/workbook/coordinator.ts";

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

void test("uses workbook:unknown coordinator key when workbook id is unavailable", async () => {
  const coordinator = new FakeCoordinator();
  const mutationEvents: WorkbookMutationEvent[] = [];
  const invalidatedWorkbookIds: Array<string | null> = [];

  const wrapped = wrapSingleTool({
    tool: makeTool("format_cells"),
    coordinator,
    contextProvider: createContextProvider(null),
    mutationEvents,
    invalidatedWorkbookIds,
  });

  await wrapped.execute("tc-null-workbook", { range: "Sheet1!A1" });

  assert.equal(coordinator.writeCalls.length, 1);
  assert.equal(coordinator.writeCalls[0]?.workbookId, "workbook:unknown");
  assert.equal(mutationEvents[0]?.workbookId, null);
  assert.equal(mutationEvents[0]?.impact, "content");
  assert.deepEqual(invalidatedWorkbookIds, []);
});
