import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { openBrowserPage, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

void test("default taskpane does not load Durable; bridge flag plus reload loads only the opt-in prototype", async () => {
  const requests: string[] = [];
  const opened = await openTaskpane(env, { prepareContext: (browser) => {
    browser.on("request", (request) => { if (request.url().includes("/src/durable/")) requests.push(request.url()); });
    return Promise.resolve();
  } });
  try {
    assert.deepEqual(requests, []);
    await Promise.all([
      opened.page.waitForResponse((response) => /\/src\/durable\/runtime\.(js|ts)/u.test(response.url())),
      opened.page.evaluate(`(async () => {
      const { executeExtensionVerificationCommand: command } = await import('/src/taskpane/background-extension-verification.ts');
      await command('setExperiment', {feature:'durable_runtime', enabled:true}, {});
      await command('reloadTaskpane', {}, {});
    })()`),
    ]);
    assert.ok(requests.some((name) => /\/src\/durable\/taskpane\.(js|ts)/u.test(name)));
    // No saved workbook exists here. Fail closed without opening writable storage.
    assert.equal((await opened.page.evaluate(() => indexedDB.databases())).some((db) => db.name?.startsWith("pi-for-excel-durable")), false);
  } finally { await opened.finish(); }
});

// A provider boundary stub, not a replacement scheduler. Durable, IndexedDB,
// fault records and tool adaptation are the production implementations.
const boot = `(async () => {
  const { isExperimentalFeatureEnabled, setExperimentalFeatureEnabled } = await import('/src/experiments/flags.ts');
  const { executeExtensionVerificationCommand: command } = await import('/src/taskpane/background-extension-verification.ts');
  await command('setExperiment', {feature: 'durable_runtime', enabled: true}, {});
  if (!isExperimentalFeatureEnabled('durable_runtime')) throw new Error('flag did not persist');
  const { openDurableRuntime } = await import('/src/durable/runtime.ts');
  const { openIndexedDbStorage } = await import('/src/durable/indexeddb-storage.ts');
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
    if (result && restarted && scenario === 'afterSync' && results.length === 1) {
      return fauxAssistantMessage(fauxToolCall(name, {}, {id: 'retry-new-id'}), {stopReason: 'toolUse'});
    }
    if (result) return fauxAssistantMessage(result.isError ? 'interrupted safely' : 'read recovered');
    return fauxAssistantMessage(fauxToolCall(name, {}, {id: 'call-one'}), {stopReason: 'toolUse'});
  };
  faux.setResponses([respond, respond, respond]);
  const models = createModels(); models.setProvider(faux.provider);
  const model = faux.getModel();
  window.runtime = await openDurableRuntime({
    storage: await openIndexedDbStorage('recovery-' + scenario), models,
    workbookId: 'test-book', model: {provider: model.provider, modelId: model.id}, allowFaultInjection: true,
    buildCapabilities: async (hooks) => ({ systemPrompt: 'Test recovery.', tools: [{
      name, label: name, description: 'Test tool', parameters: Type.Object({}),
      execute: async (id) => {
        localStorage.setItem('executions', String(Number(localStorage.getItem('executions') || 0) + 1));
        if (scenario === 'read' && !restarted) {
          localStorage.setItem('waiting-read', 'yes'); await new Promise(() => {});
        }
        if (name === 'write_cells') {
          await hooks.beforeSync({toolCallId: id, sheetId: 'sheet-id', sheetName: 'Scratch', range: 'A1',
            beforeValues: [['old']], beforeFormulas: [['']], intendedValues: [['new']]});
          localStorage.setItem('cell', 'new');
          await hooks.afterSync(id); await hooks.beforeResult(id);
        }
        return {content: [{type:'text', text:'done'}], details: {test: true}};
      }
    }]})
  });
  window.reloadViaBridge = () => command('reloadTaskpane', {}, {});
  window.ctx = ctx;
  if (!restarted) {
    if (!['read', 'model'].includes(scenario)) window.runtime.mutations.configureFault(scenario);
    const submission = await window.runtime.conversation.submit({type:'input', content:'Run the test', requestId:'original'}, ctx);
    localStorage.setItem('submission', String(submission.id));
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
    await opened.page.evaluate(`(async () => {
      const { createDurableTaskpane } = await import('/src/durable/taskpane.ts');
      const { executeExtensionVerificationCommand: command } = await import('/src/taskpane/background-extension-verification.ts');
      const { createModels } = await import('/node_modules/@earendil-works/pi-ai/dist/models.js');
      const { fauxProvider, fauxAssistantMessage, fauxToolCall } = await import('/node_modules/@earendil-works/pi-ai/dist/providers/faux.js');
      const { Type } = await import('/node_modules/typebox/build/index.mjs');
      await command('setExperiment', {feature:'durable_runtime', enabled:true}, {});
      const faux = fauxProvider({tokensPerSecond: Infinity});
      faux.setResponses([fauxAssistantMessage(fauxToolCall('write_cells', {}, {id:'bridge-call'}), {stopReason:'toolUse'}), fauxAssistantMessage('bridge complete')]);
      const models = createModels(); models.setProvider(faux.provider);
      const active = {
        agent: { state: {model:faux.getModel(), thinkingLevel:'off', isStreaming:false} },
        actionQueue: {isBusy:()=>false},
        buildDurableCapabilities: async hooks => ({systemPrompt:'Controller test', tools:[{
          name:'write_cells', description:'Write', label:'Write', parameters:Type.Object({}),
          execute: async id => {
            await hooks.beforeSync({toolCallId:id, sheetId:'sheet', sheetName:'Scratch', range:'A1', beforeValues:[['']], beforeFormulas:[['']], intendedValues:[[7]]});
            await hooks.afterSync(id); await hooks.beforeResult(id);
            return {content:[{type:'text', text:'written'}], details:{}};
          }
        }]})
      };
      const durable = await createDurableTaskpane({getActiveRuntime:()=>active, models,
        streamFn:async (...args)=>models.streamSimple(...args),
        getWorkbookContext:async()=>({workbookId:'controller-test', workbookName:'Scratch', source:'document.url'}), allowFaultInjection:true});
      window.command = (type,payload={}) => command(type,payload,{durable});
      window.waitForStatus = async predicate => {
        for (let attempt = 0; attempt < 200; attempt++) {
          if (predicate(await window.command('durableStatus'))) return;
          await new Promise(resolve => setTimeout(resolve, 50));
        }
        throw new Error('Durable status did not settle');
      };
      try { await window.command('durableFault', {point:'typo'}); throw new Error('accepted bad point'); }
      catch (e) { if (!e.message.includes('Invalid Durable')) throw e; }
      await window.command('durableFault', {point:'beforeSync'});
      await window.command('durableSubmit', {text:'Write', requestId:'bridge-request'});
    })()`);
    await opened.page.evaluate(`window.waitForStatus(status => status.fault.paused === 'beforeSync')`);
    await opened.page.evaluate(`window.command('durableResumeFault')`);
    await opened.page.evaluate(`window.waitForStatus(status => JSON.stringify(status.entries).includes('bridge complete'))`);
    const status = await opened.page.evaluate(`window.command('durableStatus')`);
    assert.deepEqual(await opened.page.evaluate(`(async () => Object.values((await window.command('durableStatus')).records).map(record => record.phase))()`), ["resultReady"], JSON.stringify(status));
  } finally { await opened.finish(); }
});

for (const scenario of ["model", "read", "beforeSync", "afterSync", "beforeResult"]) {
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
      await opened.page.evaluate(`localStorage.setItem('started', 'yes')`);
      await Promise.all([
        opened.page.waitForEvent("load"),
        opened.page.evaluate(`window.reloadViaBridge()`),
      ]);
      await opened.page.evaluate(boot);
      const result = await opened.page.evaluate<{
        settled: { status: string }; executions: number; cell: string | null;
        messages: object[]; status: { records: Record<string, { phase: string }> };
      }>(`(async () => {
        await window.runtime.conversation.waitForIdle(window.ctx);
        const status = await window.runtime.status();
        const submission = await window.runtime.harness.submission(Number(localStorage.getItem('submission')), window.ctx);
        const settled = await submission.wait(window.ctx);
        return { status, settled, executions: Number(localStorage.getItem('executions')), cell: localStorage.getItem('cell'),
          messages: (await window.runtime.conversation.context(window.ctx)).messages };
      })()`);
      assert.equal(result.settled.status, "done");
      assert.equal(result.executions, scenario === "model" ? 0 : scenario === "read" ? 2 : 1);
      assert.equal(result.cell, ["afterSync", "beforeResult"].includes(scenario) ? "new" : null);
      assert.match(JSON.stringify(result.messages), /recovered|interrupted safely/);
      if (!["model", "read"].includes(scenario)) {
        assert.equal(result.status.records["call-one"]?.phase,
          scenario === "beforeSync" ? "intended" : scenario === "afterSync" ? "synced" : "resultReady");
        assert.match(JSON.stringify(result.messages), /interrupted/);
        if (scenario === "afterSync") assert.match(JSON.stringify(result.messages), /automatic mutation retries are blocked/);
      }
    } finally { await opened.finish(); }
  });
}
