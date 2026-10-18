import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { openBrowserPage, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

for (const stored of [null, "false", "0"] as const) {
  void test(`taskpane always loads Durable with retired runtime preference ${String(stored)}`, async () => {
    const requests: string[] = [];
    const opened = await openTaskpane(env, { prepareContext: async (browser) => {
      browser.on("request", (request) => { if (request.url().includes("/src/durable/")) requests.push(request.url()); });
      await browser.addInitScript((value) => {
        if (value !== null) localStorage.setItem("pi.experimental.durableRuntime", value);
      }, stored);
    } });
    try {
      assert.ok(requests.some((name) => /\/src\/durable\/taskpane\.(js|ts)/u.test(name)));
      assert.ok(requests.some((name) => /\/src\/durable\/runtime\.(js|ts)/u.test(name)));
      assert.equal(await opened.page.evaluate(`(async () => {
        const { resolveExperimentalFeature, getExperimentalFeatureSlugs } = await import('/src/experiments/flags.ts');
        return resolveExperimentalFeature('durable-runtime') === null && !getExperimentalFeatureSlugs().includes('durable-runtime');
      })()`), true);
      const input = opened.page.locator("pi-input textarea");
      await input.fill("Chat remains available");
      assert.equal(await input.inputValue(), "Chat remains available");
      assert.equal(await input.isEnabled(), true);
    } finally { await opened.finish(); }
  });
}

// A provider boundary stub, not a replacement scheduler. Durable, IndexedDB,
// fault records and tool adaptation are the production implementations.
const boot = `(async () => {
  const { initAppStorage } = await import('/src/storage/init-app-storage.ts');
  initAppStorage();
  const { executeExtensionVerificationCommand: command } = await import('/src/taskpane/background-extension-verification.ts');
  const { createDurableTaskpane } = await import('/src/durable/taskpane.ts');
  const { createWorkbookCoordinator } = await import('/src/workbook/coordinator.ts');
  const { withWorkbookCoordinator } = await import('/src/tools/with-workbook-coordinator.ts');
  const { createWriteCellsTool } = await import('/src/tools/write-cells.ts');
  const { mutationOfficeFixture } = await import('/tests/fixtures/mutation-office.ts');
  const { createModels } = await import('/node_modules/@earendil-works/pi-ai/dist/models.js');
  const { fauxProvider, fauxAssistantMessage, fauxToolCall } = await import('/node_modules/@earendil-works/pi-ai/dist/providers/faux.js');
  const { Type } = await import('/node_modules/typebox/build/index.mjs');
  const { BACKGROUND_CONTEXT: ctx } = await import('/node_modules/@earendil-works/chord/dist/context/index.js');
  const scenario = localStorage.getItem('scenario');
  const restarted = localStorage.getItem('started') === 'yes';
  const name = scenario === 'read' ? 'read_range' : 'write_cells';
  const faux = fauxProvider({tokensPerSecond: Infinity});
  const respond = async (transcript) => {
    localStorage.setItem('requests', String(Number(localStorage.getItem('requests') || 0) + 1));
    if (scenario === 'model' && !restarted) {
      localStorage.setItem('waiting-model', 'yes');
      await new Promise(() => {});
    }
    if (scenario === 'model') return fauxAssistantMessage('recovered model');
    const results = transcript.messages.filter(m => m.role === 'toolResult');
    const result = results[0];
    if (result && restarted && scenario !== 'read') {
      const tail = JSON.stringify(transcript.messages.at(-1));
      if (!tail.includes('Do not claim these writes succeeded') || !tail.includes('do not retry writes under a new call ID')) {
        throw new Error('Restarted model request lacks authoritative mutation uncertainty guidance');
      }
    }
    if (result && restarted && scenario === 'afterDispatch' && results.length === 1) {
      return fauxAssistantMessage(fauxToolCall(name, {start_cell:'A1', values:[['new']], allow_overwrite:true}, {id: 'retry-new-id'}), {stopReason: 'toolUse'});
    }
    if (result) return fauxAssistantMessage(result.isError ? 'interrupted safely' : 'read recovered');
    return fauxAssistantMessage(fauxToolCall(name, name === 'write_cells' ? {start_cell:'A1', values:[['new']], allow_overwrite:true} : {}, {id: 'call-one'}), {stopReason: 'toolUse'});
  };
  faux.setResponses([respond, respond, respond]);
  const models = createModels(); models.setProvider(faux.provider);
  const model = faux.getModel();
  window.host = mutationOfficeFixture(JSON.parse(localStorage.getItem('workbook') || '[["old",17],["anchor",29]]'));
  window.Excel = window.host.excel;
  const coordinator = createWorkbookCoordinator();
  let active = null;
  window.durable = await createDurableTaskpane({getActiveRuntime:()=>active, models,
    streamFn:(...args)=>models.streamSimple(...args), defaultModel:model, autoCompaction:false,
    host:{kind:'office', getWorkbookContext:async()=>({workbookId:'recovery-' + scenario, workbookName:'Scratch', source:'document.url'})},
    coordinator, allowFaultInjection:true});
  window.runtime = window.durable.runtime;
  window.conversation = await window.runtime.harness.root(ctx, {agent:{model:{provider:model.provider, modelId:model.id}}});
  const write = createWriteCellsTool({appendAuditEntry:async()=>{}, appendRecoverySnapshot:async()=>null});
  const tools = [{
      name, label: name, description: 'Test tool', parameters: Type.Object({}),
      ...(name === 'write_cells' ? write : {}),
      execute: async (...args) => {
        localStorage.setItem('executions', String(Number(localStorage.getItem('executions') || 0) + 1));
        if (scenario === 'read' && !restarted) {
          localStorage.setItem('waiting-read', 'yes'); await new Promise(() => {});
        }
        if (name === 'write_cells') {
          return write.execute(...args);
        }
        return {content: [{type:'text', text:'done'}], details: {test: true}};
      }
    }];
  const capabilities = {systemPrompt:'Test recovery.', tools,
    bindMutationGate:(gate,mutationId)=>withWorkbookCoordinator(tools, coordinator,
      {getWorkbookId:async()=>window.durable.association.scopeId, getSessionId:()=>String(window.conversation.id)}, undefined,
      {mutationGate:gate, getMutationId:()=>mutationId})};
  active = {view:{conversation:window.conversation}, refreshCapabilities:()=>window.runtime.install(window.conversation, capabilities, async messages=>messages)};
  await active.refreshCapabilities();
  window.command = (type,payload={})=>command(type,payload,{durable:window.durable});
  window.reloadViaBridge = () => command('reloadTaskpane', {}, {});
  window.ctx = ctx;
  if (!restarted) {
    if (!['read', 'model'].includes(scenario)) window.runtime.mutations.configureFault(scenario);
    const submission = await window.command('durableSubmit', {text:'Run the test', requestId:'original'});
    localStorage.setItem('submission', String(submission.submissionId));
  }
})()`;

void test("verification bridge routes Durable submit, fault, status and resume through the taskpane controller", async () => {
  const opened = await openBrowserPage(env, {
    path: "durable-controller-test",
    prepareContext: async (context) => {
      await context.route("**/durable-controller-test", (route) => route.fulfill({
        contentType: "text/html", body: "<!doctype html><title>Controller test</title>",
      }));
    },
  });
  try {
    await opened.page.evaluate(`localStorage.setItem('scenario','beforeDispatch'); localStorage.setItem('started','yes')`);
    await opened.page.evaluate(boot);
    await opened.page.evaluate(`(async () => {
      window.waitForStatus = async predicate => {
        for (let attempt = 0; attempt < 200; attempt++) {
          if (predicate(await window.command('durableStatus'))) return;
          await new Promise(resolve => setTimeout(resolve, 50));
        }
        throw new Error('Durable status did not settle');
      };
      try { await window.command('durableFault', {point:'typo'}); throw new Error('accepted bad point'); }
      catch (e) { if (!e.message.includes('Invalid Durable')) throw e; }
      await window.command('durableFault', {point:'beforeDispatch'});
      await window.command('durableSubmit', {text:'Write', requestId:'bridge-request'});
    })()`);
    await opened.page.evaluate(`window.waitForStatus(status => status.fault.paused === 'beforeDispatch')`);
    const mutationId = await opened.page.evaluate(`(async () => (await window.runtime.mutations.pending())[0].mutationId)()`);
    await opened.page.evaluate(`window.command('durableResumeFault')`);
    await opened.page.evaluate(`window.conversation.waitForIdle(window.ctx)`);
    await opened.page.evaluate(`window.runtime.acknowledgeCommittedResults()`);
    const status = await opened.page.evaluate(`window.command('durableStatus')`);
    assert.equal(await opened.page.evaluate(`(async () => {
      const {MutationDoc} = await import('/src/durable/runtime.ts');
      return (await window.runtime.harness.snapshot(MutationDoc, ${JSON.stringify(mutationId)}, window.ctx)).record.phase;
    })()`), "resultCommitted", JSON.stringify(status));
    assert.deepEqual(await opened.page.evaluate(`window.runtime.mutations.pending()`), []);
    assert.equal(await opened.page.evaluate(`window.host.writes()`), 1);
  } finally { await opened.finish(); }
});

for (const scenario of ["model", "read", "beforeDispatch", "afterDispatch", "beforeResult"]) {
  void test(`reloadTaskpane recovers ${scenario} from IndexedDB without replaying a mutation`, async () => {
    const opened = await openBrowserPage(env, {
      path: "durable-recovery-test",
      prepareContext: async (context) => {
        await context.route("**/durable-recovery-test", (route) => route.fulfill({
          contentType: "text/html", body: "<!doctype html><title>Recovery test</title>",
        }));
      },
    });
    try {
      await opened.page.evaluate((value) => localStorage.setItem("scenario", value), scenario);
      await opened.page.evaluate(boot);
      await opened.page.waitForFunction(`localStorage.getItem('waiting-model') || localStorage.getItem('waiting-read') || window.runtime.mutations.status().paused`);
      assert.equal(await opened.page.evaluate(`window.host.writes()`),
        ["afterDispatch", "beforeResult"].includes(scenario) ? 1 : 0);
      const pendingBefore = await opened.page.evaluate(`(async () => (await window.runtime.mutations.pending()).map(record => record.mutationId))()`);
      await opened.page.evaluate(`localStorage.setItem('started', 'yes'); localStorage.setItem('workbook', JSON.stringify(window.host.read()))`);
      await Promise.all([
        opened.page.waitForEvent("load"),
        opened.page.evaluate(`window.reloadViaBridge()`),
      ]);
      await opened.page.evaluate(boot);
      assert.deepEqual(await opened.page.evaluate(`(async () => (await window.runtime.mutations.pending()).map(record => record.mutationId))()`), pendingBefore);
      const result = await opened.page.evaluate<{
        settled: { status: string }; executions: number; cell: string | null;
        messages: object[]; pending: object[]; writes: number; status: { records: Record<string, { phase: string }> };
      }>(`(async () => {
        await window.conversation.waitForIdle(window.ctx);
        const status = await window.command('durableStatus');
        const submission = await window.runtime.harness.submission(Number(localStorage.getItem('submission')), window.ctx);
        const settled = await submission.wait(window.ctx);
        return { status, settled, executions: Number(localStorage.getItem('executions')), cell: window.host.read()[0][0],
          writes:window.host.writes(), pending:await window.runtime.mutations.pending(), messages:status.entries };
      })()`);
      assert.equal(result.settled.status, "done");
      assert.equal(result.executions, scenario === "model" ? 0 : scenario === "read" ? 2 : scenario === "beforeDispatch" ? 0 : 1);
      assert.equal(result.cell, ["afterDispatch", "beforeResult"].includes(scenario) ? "new" : "old");
      assert.equal(result.writes, 0);
      assert.match(JSON.stringify(result.messages), /recovered|interrupted safely/);
      if (!["model", "read"].includes(scenario)) {
        assert.deepEqual(Object.values(result.status.records).map(record=>record.phase),
          [scenario === "beforeDispatch" ? "prepared" : scenario === "afterDispatch" ? "dispatched" : "returned"]);
        assert.equal(result.pending.length, 1);
        assert.match(JSON.stringify(result.messages), /interrupted/);
        if (scenario === "afterDispatch") assert.match(JSON.stringify(result.messages), /pending/);
      } else {
        assert.deepEqual(result.pending, []);
      }
    } finally { await opened.finish(); }
  });
}
