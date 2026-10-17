import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import type { MutationRecord, MutationReconciliation } from "../../src/durable/mutations.ts";
import type { JsonValue } from "@earendil-works/chord";
import { openBrowserPage, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let server: TaskpaneServer;
before(async () => { server = await startTaskpaneServer(); });
after(async () => { await server.close(); });

// Real Chromium + IndexedDB, production coordinator/controller/observer/write tool.
// The external Excel process is a queued host boundary double, not real Excel acceptance.
const boot = `(async () => {
  const { createMutationController, parseMutationRecord } = await import('/src/durable/mutations.ts');
  const { createWorkbookMutationHost } = await import('/src/durable/mutation-host.ts');
  const { createWorkbookCoordinator } = await import('/src/workbook/coordinator.ts');
  const { withWorkbookCoordinator } = await import('/src/tools/with-workbook-coordinator.ts');
  const { createWriteCellsTool } = await import('/src/tools/write-cells.ts');
  const { mutationOfficeFixture } = await import('/tests/fixtures/mutation-office.ts');
  window.host = mutationOfficeFixture(JSON.parse(localStorage.getItem('workbook') || '[["old",17],["anchor",29]]'));
  window.Excel = window.host.excel;
  const db = await new Promise((resolve, reject) => {
    const request = indexedDB.open('mutation-journal-test', 1);
    request.onupgradeneeded = () => request.result.createObjectStore('records', {keyPath:'mutationId'});
    request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error);
  });
  const transaction = (mode, work) => new Promise((resolve, reject) => {
    const tx = db.transaction('records', mode), request = work(tx.objectStore('records'));
    tx.oncomplete = () => resolve(request.result);
    tx.onerror = () => reject(tx.error); tx.onabort = () => reject(tx.error);
  });
  const store = {
    load: async id => { const value = await transaction('readonly', s => s.get(id)); return value ? parseMutationRecord(value) : null; },
    list: async () => (await transaction('readonly', s => s.getAll())).map(parseMutationRecord),
    save: record => transaction('readwrite', s => s.put(record)).then(() => undefined),
  };
  const coordinator = createWorkbookCoordinator();
  window.mutations = createMutationController({workbookId:'book', store, coordinator,
    host:createWorkbookMutationHost({kind:'office', getWorkbookId:async ()=>'book'}), allowFaultInjection:true});
  window.wrap = tool => withWorkbookCoordinator([tool], coordinator,
    {getWorkbookId:async ()=>'book', getSessionId:()=>'s'}, undefined,
    {mutationGate:window.mutations, getMutationId:id=>id})[0];
  window.tool = window.wrap(createWriteCellsTool({
    appendAuditEntry:async () => {}, appendRecoverySnapshot:async () => null
  }));
  window.write = id => window.tool.execute(id, {start_cell:'A1', values:[['new','=1+1']], allow_overwrite:true});
})()`;

for (const point of ["beforeDispatch", "afterDispatch", "beforeResult"]) {
  void test(`Chromium reload at ${point} preserves pending intent and never repeats the workbook write`, async () => {
    const opened = await openBrowserPage(server, { path: "mutation-host-test", prepareContext: async (context) => {
      await context.route("**/mutation-host-test", (route) => route.fulfill({ contentType: "text/html", body: "<!doctype html><title>Mutation test</title>" }));
    } });
    try {
      await opened.page.evaluate(boot);
      await opened.page.evaluate(`window.mutations.configureFault('${point}'); void window.write('original')`);
      await opened.page.waitForFunction(`window.mutations.status().paused === '${point}'`);
      const state = await opened.page.evaluate(`({writes:window.host.writes(), workbook:window.host.read()})`);
      assert.equal(state.writes, point === "beforeDispatch" ? 0 : 1);
      await opened.page.evaluate(`localStorage.setItem('workbook', JSON.stringify(window.host.read()))`);
      await opened.page.reload();
      await opened.page.evaluate(boot);
      const review = await opened.page.evaluate(`window.mutations.reconcile('original')`);
      assert.equal(review.classification, point === "beforeDispatch" ? "unchanged" : "applied");
      const rejected = await opened.page.evaluate<string[]>(`(async () => {
        const errors = [];
        for (const id of ['original', 'retry-new-id']) {
          try { await window.write(id); } catch (e) { errors.push(e.message); }
        }
        return errors;
      })()`);
      assert.match(rejected[0] ?? "", /replay/u);
      assert.match(rejected[1] ?? "", /pending/u);
      await opened.page.evaluate(`window.host.edit([['user edit',99],['anchor',29]])`);
      await assert.rejects(opened.page.evaluate(`window.mutations.resolve('original', '${review.observationToken}')`), /changed since review/u);
      await opened.page.evaluate(`(async () => {
        const review = await window.mutations.reconcile('original');
        await window.mutations.resolve('original', review.observationToken);
      })()`);
      assert.deepEqual(await opened.page.evaluate(`window.host.read()`), [["user edit", 99], ["anchor", 29]]);
      assert.equal(await opened.page.evaluate(`window.host.writes()`), 0);
      assert.deepEqual(await opened.page.evaluate(`window.mutations.pending()`), []);
      await opened.page.evaluate(`window.write('explicit-new-operation')`);
      assert.equal(await opened.page.evaluate(`window.host.writes()`), 1);
    } finally { await opened.finish(); }
  });
}

void test("real Office JS runner journals named-item changes but never treats the visible name as proof of all script effects", async () => {
  const opened = await openBrowserPage(server, { path: "mutation-host-test", prepareContext: async (context) => {
    await context.route("**/mutation-host-test", (route) => route.fulfill({ contentType: "text/html", body: "<!doctype html><title>Mutation test</title>" }));
  } });
  try {
    await opened.page.evaluate(boot);
    const result = await opened.page.evaluate<{
      result: JsonValue; pending: MutationRecord[]; review: MutationReconciliation;
      writes: number; names: { name: string; formula: string; type: string }[];
    }>(`(async () => {
      const {createExecuteOfficeJsTool} = await import('/src/tools/execute-office-js.ts');
      const tool = window.wrap(createExecuteOfficeJsTool());
      const result = await tool.execute('named-item', {explanation:'Create revenue name',
        code:'context.workbook.names.add("Revenue", "=Original!$A$1"); await context.sync(); return 42;'});
      return {result, pending:await window.mutations.pending(), review:await window.mutations.reconcile('named-item'),
        writes:window.host.writes(), names:window.host.names()};
    })()`);
    assert.equal(result.writes, 1);
    assert.deepEqual(result.names, [{ name: "Revenue", formula: "=Original!$A$1", type: "Range" }]);
    assert.match(JSON.stringify(result.result), /Executed Office.js/u);
    assert.equal(result.pending.length, 1);
    assert.equal(result.review.classification, "conflict");
    assert.doesNotMatch(JSON.stringify(result.pending), /context\.workbook\.names\.add/u);
    await assert.rejects(opened.page.evaluate(`window.write('new-call')`), /pending/u);
    assert.equal(await opened.page.evaluate(`window.host.writes()`), 1);
    await opened.page.evaluate(`window.mutations.resultCommitted('named-item')`);
    assert.deepEqual(await opened.page.evaluate(`window.mutations.pending()`), []);
    await opened.page.evaluate(`window.write('new-call')`);
    assert.equal(await opened.page.evaluate(`window.host.writes()`), 2);
  } finally { await opened.finish(); }
});
