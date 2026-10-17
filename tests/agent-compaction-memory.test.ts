import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { createRegistry, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { createCompactCommands } from "../src/commands/builtins/export.ts";
import { installFakeDom } from "./fixtures/fake-dom.ts";

void test("native /compact includes remembered user facts in the summary request without legacy cue rewriting", async () => {
  const dom = installFakeDom();
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  let request = "";
  models.setProvider({ ...faux.provider, streamSimple: (model, transcript, options) => {
    request = JSON.stringify(transcript);
    return faux.provider.streamSimple(model, transcript, options);
  } });
  faux.setResponses([fauxAssistantMessage("Remember EUR and calendar year")]);
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry(), settings: {
    compaction: { enabled: false, keepRecentTokens: 100 },
  } }, context);
  let view: ConversationPresenter | undefined;
  try {
    const sessions = new DurableSessions(harness, models, faux.getModel());
    const conversation = await sessions.create(faux.getModel(), "off");
    for (const fact of ["Please remember: report in EUR", "Use calendar year", "Recent context"]) {
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
    assert.match(request, /Use calendar year/u);
    assert.match(request, /focus on workbook decisions/u);
  } finally {
    await view?.dispose();
    await harness.close(context);
    dom.restore();
  }
});
