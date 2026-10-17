import assert from "node:assert/strict";
import { test } from "node:test";
import { openBrowserPage, startTaskpaneServer } from "./harness.ts";

void test("sidebar hydrates committed partials on attach and view switches and clears settled streams", async () => {
  const env = await startTaskpaneServer();
  const opened = await openBrowserPage(env, {
    path: "ui-conversation-test",
    prepareContext: async (context) => {
      await context.route("**/ui-conversation-test", (route) => route.fulfill({
        contentType: "text/html", body: "<!doctype html><title>Conversation UI</title>",
      }));
    },
  });
  try {
    const result = await opened.page.evaluate<{
      attached: string; switched: string; detached: boolean; compacting: string;
      cleared: boolean; percent: string; unsubscribed: boolean;
    }>(`(async () => {
      await import('/src/ui/register-components.ts');
      await import('/src/ui/pi-sidebar.ts');
      const { updateStatusBarForConversation } = await import('/src/taskpane/status-bar.ts');
      const model = { name: 'Projection model', id: 'projection', contextWindow: 100 };
      const message = text => ({ role: 'assistant', content: [{type: 'text', text}],
        api: 'openai-responses', provider: 'openai', model: 'projection',
        usage: {input: 40, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 40,
          cost: {input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0}},
        stopReason: 'stop', timestamp: 1 });
      // Read-side projection fixture: no execution or legacy Agent event facade.
      const projection = text => {
        const listeners = new Set();
        return { current: { messages: [], model, thinkingLevel: 'off', busy: true,
          streamingMessage: message(text), pendingToolCalls: new Set(), inbox: [],
          live: {}, providerSessionId: text },
          subscribe(listener) { listeners.add(listener); listener(this.current); return () => listeners.delete(listener); },
          publish() { for (const listener of listeners) listener(this.current); },
          listeners };
      };
      const first = projection('Already committed partial');
      const second = projection('Other conversation partial');
      const sidebar = document.createElement('pi-sidebar');
      sidebar.view = first;
      document.body.append(sidebar);
      const settle = async () => { await sidebar.updateComplete; await new Promise(requestAnimationFrame); };
      await settle();
      const attached = sidebar.querySelector('streaming-message-container').textContent;
      sidebar.view = second;
      await settle();
      const switched = sidebar.querySelector('streaming-message-container').textContent;
      const detached = first.listeners.size === 0;
      second.current = { ...second.current, busy: false, streamingMessage: undefined,
        messages: [message('Settled response'), {role: 'archivedMessages',
          archivedMessages: [{...message('Archived usage'), usage: {
            ...message('').usage, input: 90, totalTokens: 90 }}], archivedChatMessageCount: 1, timestamp: 0}],
        live: { compactions: [{taskId: 'compacting'}] } };
      second.publish();
      await settle();
      updateStatusBarForConversation(second);
      const compacting = sidebar.querySelector('pi-working-indicator').primaryText;
      const cleared = sidebar.querySelector('streaming-message-container').textContent.trim() === '';
      const percent = sidebar.querySelector('.pi-status-ctx__pct').textContent;
      sidebar.remove();
      return { attached, switched, detached, compacting, cleared, percent, unsubscribed: second.listeners.size === 0 };
    })()`);
    assert.match(result.attached, /Already committed partial/);
    assert.match(result.switched, /Other conversation partial/);
    assert.equal(result.detached, true);
    assert.equal(result.compacting, "Compacting context…");
    assert.equal(result.cleared, true);
    assert.equal(result.percent, "40%");
    assert.equal(result.unsubscribed, true);
  } finally {
    await opened.finish();
    await env.close();
  }
});
