import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

void test("a slow check of the previous proxy URL does not overwrite the check of the current one", async () => {
  const opened = await openTaskpane(env, { clientId: "proxy-status-browser" });
  const { page } = opened;
  try {
    await page.route("http://localhost:4901/healthz", async (route) => {
      await new Promise((resolve) => setTimeout(resolve, 600));
      await route.fulfill({ status: 200, headers: { "access-control-allow-origin": "*" }, body: "ok" });
    });
    await page.route("http://localhost:4902/healthz", (route) => route.abort("connectionrefused"));

    const state = await page.evaluate(async () => {
      const path = "/src/taskpane/proxy-status.ts";
      const { checkProxyOnce, getProxyState } = await import(/* @vite-ignore */ path) as typeof import("../../src/taskpane/proxy-status.ts");
      let url = "http://localhost:4901";
      const settings = { get: (key: string) => Promise.resolve(key === "proxy.url" ? url : null) };
      const previous = checkProxyOnce(settings);
      await new Promise((resolve) => setTimeout(resolve, 50));
      url = "http://localhost:4902";
      const current = await checkProxyOnce(settings, { force: true });
      await previous;
      return { current, published: getProxyState() };
    });

    assert.deepEqual(state, { current: "not-detected", published: "not-detected" });
  } finally {
    await opened.finish();
  }
});
