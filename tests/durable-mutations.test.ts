import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { JsonValue } from "@earendil-works/chord";
import { createMutationController, parseMutationRecord, MutationExecutionError, type MutationRecord, type MutationObservation,
  type MutationHost, type MutationFaultPoint } from "../src/durable/mutations.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";
import { getToolExecutionMode } from "../src/tools/execution-policy.ts";
import { CORE_TOOL_NAMES } from "../src/tools/names.ts";
import { createWorkbookMutationHost } from "../src/durable/mutation-host.ts";

function fixture(incomplete = false) {
  const records = new Map<string, MutationRecord>();
  const observation: MutationObservation = {
    workbookId: "book", structure: [{ id: "sheet-id", tables: [{ id: "table-id", range: "B2:D9" }] }],
    state: [["old", 17]], complete: !incomplete, limitation: incomplete ? "Unbounded effects" : null,
  };
  let effects = 0;
  let failSave: MutationRecord["phase"] | null = null;
  let failList = false;
  const store = {
    load: (id: string) => Promise.resolve(structuredClone(records.get(id) ?? null)),
    list: () => {
      if (failList) return Promise.reject(new Error("list unavailable"));
      return Promise.resolve(structuredClone([...records.values()]));
    },
    save: (record: MutationRecord) => {
      if (failSave === record.phase) return Promise.reject(new Error("commit unavailable"));
      records.set(record.mutationId, structuredClone(record));
      return Promise.resolve();
    },
  };
  const host: MutationHost = {
    prepare: (call) => Promise.resolve({ intent: call.params, target: { sheetId: "sheet-id" },
      before: structuredClone(observation), expected: { ...structuredClone(observation), state: [["new", 23]] } }),
    observe: () => Promise.resolve(structuredClone(observation)),
  };
  const coordinator = createWorkbookCoordinator();
  const open = () => createMutationController({ workbookId: "book", store, host, coordinator, allowFaultInjection: true });
  const controller = open();
  const tool = (name = "write_cells", failure = false, gate = controller,
    receipt: AgentToolResult<JsonValue> = { content: [{ type: "text", text: "host returned" }], details: {} },
    getMutationId = (id: string): string => id) => {
    const original: AgentTool = { name, label: name, description: name, parameters: Type.Object({}),
      execute: () => {
        // Independently assert the write-ahead invariant at the external side-effect boundary.
        assert.ok([...records.values()].some((record) => record.phase === "dispatched"));
        effects += 1;
        observation.state = [["new", 23]];
        if (failure) return Promise.reject(new Error("host disconnected after applying change"));
        return Promise.resolve(receipt);
      } };
    const wrapped = withWorkbookCoordinator([original], coordinator,
      { getWorkbookId: () => Promise.resolve("book"), getSessionId: () => "session" }, undefined,
      { mutationGate: gate, getMutationId })[0];
    assert.ok(wrapped);
    return wrapped;
  };
  return { records, store, host, coordinator, controller, open, tool,
    effects: () => effects, edit: (state: JsonValue) => { observation.state = state; },
    structure: (state: JsonValue) => { observation.structure = state; },
    failSave: (phase: MutationRecord["phase"] | null) => { failSave = phase; },
    failList: () => { failList = true; },
  };
}

for (const point of ["beforeDispatch", "afterDispatch", "beforeResult"] satisfies MutationFaultPoint[]) {
  void test(`committed journal survives ${point}; reload cannot replay and resolution never writes`, async () => {
    const f = fixture();
    f.controller.configureFault(point);
    const execution = f.tool().execute("original", {});
    for (let i = 0; i < 100 && !f.controller.status().paused; i++) await new Promise<void>((resolve) => setImmediate(resolve));
    assert.equal(f.controller.status().paused, point);
    assert.equal(f.effects(), point === "beforeDispatch" ? 0 : 1);
    const recovered = f.open();
    const review = await recovered.reconcile("original");
    assert.equal(review.classification, point === "beforeDispatch" ? "unchanged" : "applied");
    assert.equal((await recovered.pending()).length, 1);
    f.controller.resume();
    await execution;
    await assert.rejects(f.tool("write_cells", false, recovered).execute("original", {}), /replay/u);
    await assert.rejects(f.tool("format_cells", false, recovered).execute("new-call", {}), /pending/u);
    const reviewed = await recovered.reconcile("original");
    f.edit([["user edit", 99]]);
    await assert.rejects(recovered.resolve("original", reviewed.observationToken), /changed since review/u);
    const conflict = await recovered.reconcile("original");
    assert.equal(conflict.classification, "conflict");
    const beforeResolution = f.effects();
    const revision = f.coordinator.getRevision("book");
    await recovered.resolve("original", conflict.observationToken);
    assert.equal(f.coordinator.getRevision("book"), revision);
    assert.equal(f.effects(), beforeResolution);
    assert.deepEqual((await recovered.reconcile("original")).observation.state, [["user edit", 99]]);
    assert.deepEqual(await recovered.pending(), []);
    await f.tool("format_cells", false, recovered).execute("new-call", {});
    assert.equal(f.effects(), beforeResolution + 1);
    await recovered.resultCommitted("new-call");
    assert.deepEqual(await recovered.pending(), []);
    assert.equal(f.coordinator.getRevision("book"), revision + 1);
  });
}

// This is an execution-policy coverage matrix, not a second production tool registry.
const mutationCases: [string, JsonValue][] = CORE_TOOL_NAMES
  .map((name): [string, JsonValue] => [name, {}])
  .filter(([name, params]) => getToolExecutionMode(name, params) === "mutate");
for (const action of ["insert_rows", "delete_rows", "insert_columns", "delete_columns", "add_sheet", "delete_sheet", "rename_sheet", "duplicate_sheet", "hide_sheet", "unhide_sheet"]) {
  mutationCases.push(["modify_structure", { action }]);
}
for (const action of ["add", "update", "reply", "delete", "resolve", "reopen"]) mutationCases.push(["comments", { action }]);
for (const action of ["create", "update", "delete"]) mutationCases.push(["charts", { action }]);
mutationCases.push(["workbook_history", { action: "restore" }], ["python_transform_range", {}],
  ["execute_office_js", { code: "workbook.names.add(...)" }], ["execute_wps_js", {}], ["extension_workbook_change", {}]);

for (const [name, params] of mutationCases) {
  void test(`${name} ${JSON.stringify(params)}: host-side interruption cannot duplicate effects`, async () => {
    const f = fixture(true);
    await assert.rejects(f.tool(name, true).execute("interrupted", params), /may already have applied/u);
    assert.equal(f.effects(), 1);
    const restarted = f.open();
    assert.equal((await restarted.reconcile("interrupted")).classification, "conflict");
    await assert.rejects(f.tool(name, false, restarted).execute("interrupted", params), /replay/u);
    await assert.rejects(f.tool(name, false, restarted).execute("retry-with-new-id", params), /pending/u);
    assert.equal(f.effects(), 1);
  });
}

void test("failure before dispatch is nonmutating; uncertain commit/read failures fail closed", async () => {
  for (const phase of ["prepared", "dispatched", "returned"] satisfies MutationRecord["phase"][]) {
    const f = fixture();
    f.failSave(phase);
    await assert.rejects(f.tool().execute("original", {}), (error: Error) => {
      assert.ok(error instanceof MutationExecutionError);
      assert.equal(error.mutationId, "original");
      assert.equal(error.outcome, phase === "returned" ? "uncertain" : "not-executed");
      assert.match(error.message, phase === "returned" ? /may already have applied/u : /not executed/u);
      assert.doesNotMatch(error.message, /commit unavailable/u);
      return true;
    });
    assert.equal(f.effects(), phase === "returned" ? 1 : 0);
    f.failSave(null);
    if (phase !== "prepared") await assert.rejects(f.tool().execute("retry", {}), /pending/u);
  }
  const f = fixture();
  f.failList();
  await assert.rejects(f.tool().execute("call", {}), /list unavailable/u);
  assert.equal(f.effects(), 0);
});

for (const failure of ["returned-save", "post-host-observation"]) {
  void test(`${failure}: surface safe uncertainty and retain dispatched intent after host changes`, async () => {
    const f = fixture();
    const secret = "private-storage-url-and-token";
    const save = f.store.save;
    const observe = f.host.observe;
    f.store.save = (record) => failure === "returned-save" && record.phase === "returned"
      ? Promise.reject(new Error(secret, { cause: new Error(secret) })) : save(record);
    f.host.observe = (record) => failure === "post-host-observation" && f.effects() > 0
      ? Promise.reject(new Error(secret, { cause: new Error(secret) })) : observe(record);
    await assert.rejects(f.tool().execute("uncertain", {}), (error: Error) => {
      assert.ok(error instanceof MutationExecutionError);
      assert.equal(error.outcome, "uncertain");
      assert.equal(error.mutationId, "uncertain");
      assert.match(error.message, /may already have applied/u);
      assert.match(error.message, /review/iu);
      assert.match(error.message, /do not retry automatically/iu);
      assert.doesNotMatch(`${error.message} ${error.stack ?? ""} ${JSON.stringify(error)}`, /private-storage-url-and-token/u);
      assert.equal(error.cause, undefined);
      return true;
    });
    assert.equal(f.effects(), 1);
    const record = f.records.get("uncertain");
    assert.ok(record);
    assert.equal(record.phase, "dispatched");
    assert.deepEqual((await observe(record)).state, [["new", 23]]);
    const recovered = f.open();
    assert.equal((await recovered.pending()).length, 1);
    await assert.rejects(f.tool("write_cells", false, recovered).execute("next", {}), /pending/u);
    await assert.rejects(f.tool("write_cells", false, recovered).execute("uncertain", {}), /replay/u);
    assert.equal(f.effects(), 1);
  });
}

for (const [kind, name] of [["browser", "extension_custom"], ["wps", "execute_wps_js"]] as const) {
  void test(`${name}: two successful committed results proceed despite incomplete observations`, async () => {
    const f = fixture(true);
    Object.assign(f.host, createWorkbookMutationHost({ kind, getWorkbookId: () => Promise.resolve("book") }));
    for (const id of ["first", "second"]) {
      await f.tool(name).execute(id, {});
      assert.equal((await f.controller.reconcile(id)).classification, "conflict");
      assert.equal((await f.controller.pending()).length, 1);
      await assert.rejects(f.tool(name).execute("too-early", {}), /pending/u);
      await f.controller.resultCommitted(id);
      assert.deepEqual(await f.open().pending(), []);
      await assert.rejects(f.tool(name).execute(id, {}), /replay/u);
    }
    assert.equal(f.effects(), 2);
  });
}

void test("failed receipts stay pending after commit, including structured errors without error text", async () => {
  const receipts: AgentToolResult<JsonValue>[] = [
    { content: [{ type: "text", text: "Error executing WPS JSAPI: disconnected" }], details: {} },
    { content: [{ type: "text", text: "Host returned" }], details: { kind: "python_transform_range", blocked: false, error: "Partial write" } },
    { content: [], details: { kind: "mcp_gateway", operation: "call_tool", ok: false } },
    { content: [], details: { kind: "workbook_history", action: "restore", error: "Restore incomplete" } },
  ];
  for (const receipt of receipts) {
    const f = fixture(true);
    await f.tool("extension_custom", false, f.controller, receipt).execute("failed", {});
    await f.controller.resultCommitted("failed");
    assert.equal((await f.open().pending()).length, 1);
    await assert.rejects(f.tool().execute("next", {}), /pending/u);
    const review = await f.controller.reconcile("failed");
    await f.controller.resolve("failed", review.observationToken);
    assert.equal(f.effects(), 1);
  }
});

void test("logical task IDs distinguish reused provider IDs while blocking same-task replay", async () => {
  const f = fixture();
  for (const taskId of ["conversation-a/task-1", "conversation-a/task-2", "conversation-b/task-1"]) {
    const tool = f.tool("write_cells", false, f.controller, undefined, () => taskId);
    await tool.execute("provider-reused-id", {});
    await f.controller.resultCommitted(taskId);
    assert.equal(f.records.get(taskId)?.toolCallId, "provider-reused-id");
    await assert.rejects(tool.execute("provider-reused-id", {}), /replay/u);
  }
  assert.equal(f.effects(), 3);
  assert.equal(f.records.size, 3);
});

void test("structural identity changes invalidate matching cell contents", async () => {
  const f = fixture();
  await f.tool().execute("cells", {});
  f.structure([{ id: "replacement-sheet-id", tables: [{ id: "replacement-table", range: "B2:D9" }] }]);
  assert.equal((await f.controller.reconcile("cells")).classification, "conflict");
});

void test("malformed or legacy journal records cannot become absent or completed mutations", () => {
  assert.throws(() => parseMutationRecord({ version: 1, phase: "resultReady" }), /Invalid/u);
  assert.throws(() => parseMutationRecord({ version: 2, phase: "resolved" }), /Invalid/u);
});

void test("production generic/WPS intent omits executable code and cannot claim applied/unchanged", async () => {
  const host = createWorkbookMutationHost({ kind: "wps", getWorkbookId: () => Promise.resolve("book") });
  const plan = await host.prepare({ mutationId: "task/generic", toolCallId: "generic", toolName: "execute_wps_js", workbookId: "book", sessionId: "s",
    params: { code: "sensitive arbitrary script", password: "sensitive argument" } });
  assert.equal(plan.before.complete, false);
  assert.doesNotMatch(JSON.stringify(plan), /sensitive/u);
});

for (const incomplete of [false, true]) {
  void test(`prepared intent reports conflict for user edits or incomplete observations (${incomplete})`, async () => {
    const f = fixture(incomplete);
    f.failSave("dispatched");
    await assert.rejects(f.tool().execute("prepared", {}), /not executed/u);
    f.failSave(null);
    if (!incomplete) f.edit([["user edit", 99]]);
    const review = await f.controller.reconcile("prepared");
    assert.equal(review.classification, "conflict");
    assert.match(review.reason, /never entered/u);
    await f.controller.resolve("prepared", review.observationToken);
    assert.equal(f.effects(), 0);
    assert.deepEqual((await f.controller.reconcile("prepared")).observation.state,
      incomplete ? [["old", 17]] : [["user edit", 99]]);
  });
}
