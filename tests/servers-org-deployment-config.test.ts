import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createModels,
  fauxAssistantMessage,
  fauxProvider,
  type Model,
  type StreamOptions,
} from "@earendil-works/pi-ai";

import { createOfficeStreamFn } from "../src/auth/stream-proxy.ts";
import {
  CODEX_WEBSOCKET_BRIDGE_HEADER,
  PROXY_HEALTH_HEADER,
} from "../src/auth/proxy-validation.ts";

async function withBrowserBoundaries<T>(run: () => Promise<T>): Promise<T> {
  const previousFetch = globalThis.fetch;
  const hadDocument = Reflect.has(globalThis, "document");
  const previousDocument = Reflect.get(globalThis, "document");
  Reflect.set(globalThis, "document", new EventTarget());
  Reflect.set(globalThis, "fetch", () => Promise.resolve(new Response("ok", {
    status: 200,
    headers: {
      [PROXY_HEALTH_HEADER]: "1",
      [CODEX_WEBSOCKET_BRIDGE_HEADER]: "1",
    },
  })));
  try {
    return await run();
  } finally {
    Reflect.set(globalThis, "fetch", previousFetch);
    if (hadDocument) Reflect.set(globalThis, "document", previousDocument);
    else Reflect.deleteProperty(globalThis, "document");
  }
}

void test("connection API selects the Codex WebSocket bridge only for ChatGPT GPT-5.6 Luna", async () => {
  await withBrowserBoundaries(async () => {
    const rows = [
      { provider: "openai-codex", id: "gpt-5.6-luna", expectedTransport: true },
      { provider: "openai-codex", id: "gpt-5.6-sol", expectedTransport: false },
      { provider: "openai", id: "gpt-5.6-luna", expectedTransport: false },
    ] as const;

    for (const [index, row] of rows.entries()) {
      const faux = fauxProvider({
        provider: row.provider,
        api: "openai-codex-responses",
        models: [{ id: row.id }],
      });
      const models = createModels();
      models.setProvider(faux.provider);
      let streamedModel: Model<string> | undefined;
      faux.setResponses([(_context, _options, _state, model) => {
        streamedModel = model;
        return fauxAssistantMessage("done");
      }]);

      const stream = await createOfficeStreamFn(
        () => Promise.resolve(`https://proxy-${index}.example`),
        models,
      )(faux.getModel(), { messages: [] });
      await stream.result();

      assert.equal(streamedModel?.baseUrl.includes("pi_transport=codex-websocket"), row.expectedTransport);
    }
  });
});

void test("connection API preserves UUIDv7 and stably maps legacy bridge session ids", async () => {
  await withBrowserBoundaries(async () => {
    const faux = fauxProvider({
      provider: "openai-codex",
      api: "openai-codex-responses",
      models: [{ id: "gpt-5.6-luna" }],
    });
    const models = createModels();
    models.setProvider(faux.provider);
    const seenSessionIds: Array<string | undefined> = [];
    faux.setResponses(Array.from({ length: 3 }, () => (
      _context: unknown,
      options: StreamOptions | undefined,
    ) => {
      seenSessionIds.push(options?.sessionId);
      return fauxAssistantMessage("done");
    }));
    const stream = createOfficeStreamFn(() => Promise.resolve("https://session-proxy.example"), models);
    const nativeSessionId = "019f4c1c-03ae-7d15-8e28-035d6a58c787";
    const legacySessionId = "b01d800c-e36c-4737-b987-c5ebb16d4106";

    for (const sessionId of [nativeSessionId, legacySessionId, legacySessionId]) {
      const response = await stream(faux.getModel(), { messages: [] }, { sessionId });
      await response.result();
    }

    assert.equal(seenSessionIds[0], nativeSessionId);
    assert.match(seenSessionIds[1] ?? "", /^[0-9a-f]{8}-[0-9a-f]{4}-7[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/u);
    assert.equal(seenSessionIds[2], seenSessionIds[1]);
  });
});
