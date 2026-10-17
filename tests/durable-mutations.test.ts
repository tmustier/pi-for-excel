import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { JsonValue } from "@earendil-works/chord";
import { createMutationController, parseMutationRecord, type MutationRecord, type MutationObservation,
  type MutationHost, type MutationFaultPoint } from "../src/durable/mutations.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";
import { getToolExecutionMode } from "../src/tools/execution-policy.ts";
import { CORE_TOOL_NAMES } from "../src/tools/names.ts";
import { createWorkbookMutationHost } from "../src/durable/mutation-host.ts";

function fixture(requiresReview = false) {
  const records = new Map<string, MutationRecord>();
  const observation: MutationObservation = {
    workbookId: "book", structure: [{ id: "sheet-id", tables: [{ id: "table-id", range: "B2:D9" }] }],
    state: [["old", 17]], complete: !requiresReview, limitation: requiresReview ? "Unbounded effects" : null,
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
      records.set(record.toolCallId, structuredClone(record));
      return Promise.resolve();
    },
  };
  const host: MutationHost = {
    prepare: (call) => Promise.resolve({ intent: call.params, target: { sheetId: "sheet-id" },
      before: structuredClone(observation), expected: { ...structuredClone(observation), state: [["new", 23]] }, requiresReview }),
    observe: () => Promise.resolve(structuredClone(observation)),
  };
  const coordinator = createWorkbookCoordinator();
  const open = () => createMutationController({ workbookId: "book", store, host, coordinator, allowFaultInjection: true });
  const controller = open();
  const tool = (name = "write_cells", failure = false, gate = controller) => {
    const original: AgentTool = { name, label: name, description: name, parameters: Type.Object({}),
      execute: () => {
        // Independently assert the write-ahead invariant at the external side-effect boundary.
        assert.ok([...records.values()].some((record) => record.phase === "dispatched"));
        effects += 1;
        observation.state = [["new", 23]];
        if (failure) return Promise.reject(new Error("host disconnected after applying change"));
        return Promise.resolve({ content: [{ type: "text", text: "host returned" }], details: {} });
      } };
    const wrapped = withWorkbookCoordinator([original], coordinator,
      { getWorkbookId: () => Promise.resolve("book"), getSessionId: () => "session" }, undefined,
      { mutationGate: gate })[0];
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
    await assert.rejects(f.tool(name, true).execute("interrupted", params), /disconnected/u);
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
    await assert.rejects(f.tool().execute("original", {}), /commit unavailable/u);
    assert.equal(f.effects(), phase === "returned" ? 1 : 0);
    f.failSave(null);
    if (phase !== "prepared") await assert.rejects(f.tool().execute("retry", {}), /pending/u);
  }
  const f = fixture();
  f.failList();
  await assert.rejects(f.tool().execute("call", {}), /list unavailable/u);
  assert.equal(f.effects(), 0);
});

void test("unknown effects remain pending after result commit, and incomplete equal snapshots never prove success", async () => {
  const f = fixture(true);
  await f.tool("extension_custom").execute("generic", {});
  await f.controller.resultCommitted("generic");
  assert.equal((await f.controller.pending()).length, 1);
  assert.equal((await f.controller.reconcile("generic")).classification, "conflict");
  const review = await f.controller.reconcile("generic");
  await f.controller.resolve("generic", review.observationToken);
  assert.equal(f.effects(), 1);
  assert.equal((await f.controller.pending()).length, 0);
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
  const plan = await host.prepare({ toolCallId: "generic", toolName: "execute_wps_js", workbookId: "book", sessionId: "s",
    params: { code: "sensitive arbitrary script", password: "sensitive argument" } });
  assert.equal(plan.before.complete, false);
  assert.equal(plan.requiresReview, true);
  assert.doesNotMatch(JSON.stringify(plan), /sensitive/u);
});
