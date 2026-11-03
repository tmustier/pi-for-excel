import { authDeferred } from "./fixtures/auth-deferred.ts";
import assert from "node:assert/strict";
import { test } from "node:test";
import { createModels } from "@earendil-works/pi-ai";
import { getBuiltinModels } from "@earendil-works/pi-ai/providers/all";

import { createOfficeStreamFn } from "../src/auth/stream-proxy.ts";

void test("latest-started Codex capability probe owns the cache even when an older probe finishes last", async (t) => {
  const previousDocument = Reflect.get(globalThis, "document");
  Reflect.set(globalThis, "document", new EventTarget());
  const previousFetch = globalThis.fetch;
  const pending: Array<ReturnType<typeof authDeferred<Response>>> = [];
  globalThis.fetch = () => {
    const response = authDeferred<Response>();
    pending.push(response);
    return response.promise;
  };
  t.after(() => {
    globalThis.fetch = previousFetch;
    Reflect.set(globalThis, "document", previousDocument);
  });
  const baseModel = getBuiltinModels("openai-codex")[0];
  assert.ok(baseModel);
  const model = { ...baseModel, id: "gpt-5.6-luna" };
  const stream = createOfficeStreamFn(() => Promise.resolve("https://localhost:3099"), createModels());
  const older = stream(model, { messages: [] });
  const newer = stream(model, { messages: [] });
  await Promise.resolve();
  assert.equal(pending.length, 2);
  const [oldProbe, newProbe] = pending;
  assert.ok(oldProbe && newProbe);
  newProbe.resolve(new Response("unsupported", { status: 404 }));
  await assert.rejects(newer, /newer Pi for Excel proxy/);
  oldProbe.resolve(new Response("ok", {
    headers: { "x-pi-for-excel-proxy": "1", "x-pi-for-excel-codex-websocket-bridge": "1" },
  }));
  await (await older).result();
  await assert.rejects(stream(model, { messages: [] }), /newer Pi for Excel proxy/);
  assert.equal(pending.length, 2, "cached newest result must be reused");
});
