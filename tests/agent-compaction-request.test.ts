// Command contracts use a real native Harness; the DOM double only captures toasts.
import assert from "node:assert/strict";
import { after, test } from "node:test";
import { copyJson } from "@earendil-works/chord";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { CompactionEntry, createRegistry, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { ConversationPresenter, readConversationHistory } from "../src/durable/conversation-view.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { createCompactCommands, createExportCommands } from "../src/commands/builtins/export.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

// The shared toast component retains its DOM node across command invocations.
const dom = installFakeDom();
after(() => dom.restore());

void test("/export includes reset-hidden immutable history and catalogue metadata", async () => {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  models.setProvider(faux.provider);
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  let view: ConversationPresenter | undefined;
  const previousNavigator = Object.getOwnPropertyDescriptor(globalThis, "navigator");
  let copied = "";
  Object.defineProperty(globalThis, "navigator", { configurable: true, value: {
    clipboard: { writeText: (value: string) => { copied = value; return Promise.resolve(); } },
  } });
  try {
    const sessions = new DurableSessions(harness, models, faux.getModel());
    const conversation = await sessions.create(faux.getModel(), "off", "Native export");
    await conversation.submit({ type: "write", entry: { kind: "test.history", model: [
      { role: "user", content: "hidden original", timestamp: 1000 },
    ] } }, context);
    await conversation.reset(undefined, context);
    view = await ConversationPresenter.open(harness, conversation, models, faux.getModel());
    const active = view;
    const command = createExportCommands(() => active, () => sessions)[0];
    assert.ok(command);
    await command.execute("clipboard");
    const expected = await sessions.loadSession(String(conversation.id));
    assert.ok(expected);
    const exported = copyJson(JSON.parse(copied));
    assert.ok(exported && typeof exported === "object" && !Array.isArray(exported));
    assert.deepEqual(exported.raw, expected.messages);
    assert.equal(exported.title, "Native export");
    assert.equal(exported.id, String(conversation.id));
    assert.deepEqual(exported.transcript, [{ role: "user", text: "hidden original" }]);
  } finally {
    await view?.dispose();
    await harness.close(context);
    if (previousNavigator) Object.defineProperty(globalThis, "navigator", previousNavigator);
    else Reflect.deleteProperty(globalThis, "navigator");
  }
});

for (const failure of [false, true]) {
  void test(`/compact waits for native ${failure ? "failure" : "summary"} and forwards focus`, async () => {
    const models = createModels();
    const faux = fauxProvider({ tokensPerSecond: Infinity });
    let prompt = "";
    models.setProvider({ ...faux.provider, streamSimple: (model, transcript, options) => {
      prompt = JSON.stringify(transcript);
      return faux.provider.streamSimple(model, transcript, options);
    } });
    faux.setResponses([failure
      ? fauxAssistantMessage("", { stopReason: "error", errorMessage: "native summary failed" })
      : fauxAssistantMessage("native summary")]);
    const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry(), settings: {
      compaction: { enabled: false, keepRecentTokens: 100 },
    } }, context);
    let view: ConversationPresenter | undefined;
    try {
      const sessions = new DurableSessions(harness, models, faux.getModel());
      const conversation = await sessions.create(faux.getModel(), "off");
      view = await ConversationPresenter.open(harness, conversation, models, faux.getModel());
      const active = view;
      const command = createCompactCommands(() => active)[0];
      assert.ok(command);
      await command.execute("");
      assert.equal(prompt, "", "empty native context needs no summary request");
      assert.match(dom.document.getElementById("pi-toast")?.children[0]?.children[0]?.textContent ?? "", /Nothing to compact/u);
      for (let i = 0; i < 6; i++) {
        await conversation.submit({ type: "write", entry: { kind: "test.history", model: [
          { role: "user", content: "history ".repeat(500), timestamp: i },
        ] } }, context);
      }
      await command.execute("  preserve formulas  ");
      assert.match(prompt, /preserve formulas/u);
      const summaries = (await readConversationHistory(conversation)).filter(CompactionEntry.is);
      assert.equal(summaries.length, failure ? 0 : 1);
      const toast = dom.document.getElementById("pi-toast")?.children[0]?.children[0]?.textContent ?? "";
      assert.match(toast, failure ? /native summary failed/u : /✓/u);
    } finally {
      await view?.dispose();
      await harness.close(context);
    }
  });
}
