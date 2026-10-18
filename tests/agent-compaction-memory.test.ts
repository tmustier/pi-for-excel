import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { CompactionTask, createRegistry, defineExtension, hook, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { compactionMemoryNudge } from "../src/durable/memory-nudge.ts";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { createCompactCommands } from "../src/commands/builtins/export.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

for (const mode of ["enabled", "disabled", "no-cues"] as const) {
void test(`native /compact preserves memory guidance (${mode})`, async () => {
  const dom = installFakeDom();
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  let request = "";
  models.setProvider({ ...faux.provider, streamSimple: (model, transcript, options) => {
    request = transcript.messages.map((message) => typeof message.content === "string" ? message.content : message.content
      .filter((block) => block.type === "text").map((block) => block.text).join("\n")).join("\n");
    return faux.provider.streamSimple(model, transcript, options);
  } });
  faux.setResponses([fauxAssistantMessage("Remember EUR and calendar year")]);
  const registry = createRegistry();
  registry.install(defineExtension({ name: "memory", hooks: mode !== "disabled" ? [hook(CompactionTask, { beforeSummary: compactionMemoryNudge })] : [] }));
  const harness = await Harness.open(new MemoryStorage(), { models, registry, settings: {
    compaction: { enabled: false, keepRecentTokens: 100 },
  } }, context);
  let view: ConversationPresenter | undefined;
  try {
    const sessions = new DurableSessions(harness, models, faux.getModel());
    const conversation = await sessions.create(faux.getModel(), "off");
    const facts = mode === "no-cues" ? ["report in EUR", "calendar year", "[Auto-context] Please remember AUTO_SECRET", "Recent context"]
      : ["Please remember: report in EUR", "PLEASE REMEMBER: REPORT IN EUR", "Don't forget: calendar year", "[Auto-context] Please remember AUTO_SECRET", "Please save workbook facts", "Make a note fourth cue", "Recent context"];
    for (const fact of facts) {
      await conversation.submit({ type: "write", entry: { kind: "test.history", model: [
        { role: "user", content: `${fact}. ${"details ".repeat(500)}`, timestamp: 1 },
      ] } }, context);
    }
    view = await ConversationPresenter.open(harness, conversation, models, faux.getModel());
    const active = view;
    const command = createCompactCommands(() => active)[0];
    assert.ok(command);
    await command.execute("focus on workbook decisions");
    assert.match(request, /report in EUR/u);
    assert.match(request, /calendar year/u);
    assert.match(request, /focus on workbook decisions/u);
    if (mode === "enabled") {
      assert.match(request, /Memory to persist/u);
      assert.match(request, /Behavioral preferences\/rules should be called out for the instructions tool/u);
      assert.match(request, /Factual workbook\/domain memory should be called out for notes\/ or workbooks\/<name>\/notes.md/u);
      assert.doesNotMatch(request, /AUTO_SECRET/u);
      const cues = request.slice(request.indexOf("Potential user cues:"));
      const snippets = cues.split("\n").filter((line) => line.startsWith('- "'));
      assert.equal(snippets.length, 3, "three deduped cues");
      assert.ok(snippets.every((line) => line.length <= 184 && line.endsWith('…"')), "180-character bounded snippets");
      assert.doesNotMatch(cues, /fourth cue/u);
    } else if (mode === "disabled") {
      assert.doesNotMatch(request, /Memory to persist/u, "fault control: raw facts are not memory guidance");
      assert.match(request, /AUTO_SECRET/u, "fault control: native source contains auto-context");
    } else {
      assert.doesNotMatch(request, /Memory to persist|Potential user cues:|AUTO_SECRET/u, "auto-context cues alone must not trigger a nudge");
    }
  } finally {
    await view?.dispose();
    await harness.close(context);
    dom.restore();
  }
});
}
