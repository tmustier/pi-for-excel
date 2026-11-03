import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { enterCommand, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

async function chooseModel(page: Page, search: string, provider: string): Promise<void> {
  await enterCommand(page, "/model");
  const picker = page.locator("#pi-model-selector-overlay");
  await picker.locator(".pi-model-selector-search").fill(search);
  await picker.locator(".pi-model-selector-item")
    .filter({ has: page.locator(".pi-model-selector-item-provider", { hasText: new RegExp(`^${provider}$`, "u") }) })
    .first().click();
  await picker.waitFor({ state: "detached" });
}

void test("a login that expires in a stopped turn does not label the next turn's unrelated failure", async () => {
  const opened = await openTaskpane(env, { viewport: { width: 420, height: 720 }, clientId: "turn-errors-browser" });
  const { page } = opened;
  try {
    let releaseRefresh: () => void = () => {};
    const refreshHeld = new Promise<void>((resolve) => {
      releaseRefresh = resolve;
    });
    await page.route("**/v1/oauth/token", async (route) => {
      await refreshHeld;
      await route.fulfill({
        status: 400,
        contentType: "application/json",
        headers: { "access-control-allow-origin": "*" },
        body: JSON.stringify({ error: "invalid_grant" }),
      });
    });
    await page.evaluate(async () => {
      const storage = await import("../../src/storage/local/app-storage.ts");
      const appStorage = storage.getAppStorage();
      await appStorage.settings.set("oauth.anthropic", {
        access: "stale-access", refresh: "dead-refresh", expires: Date.now() - 60_000,
      });
      await appStorage.providerKeys.set("anthropic", "stale-access");
      await appStorage.providerKeys.set("mistral", "fixture-key");
    });
    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });

    await chooseModel(page, "claude", "anthropic");
    await enterCommand(page, "Hello");
    const stop = page.locator(".pi-input-btn--abort");
    await stop.click();
    releaseRefresh();
    await stop.waitFor({ state: "detached" });

    await chooseModel(page, "mistral", "mistral");
    await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar)) throw new Error("sidebar missing");
      const agent = sidebar.agent as { streamFunction: () => Promise<never> };
      agent.streamFunction = () => Promise.reject(new Error("Bad request: unsupported parameter"));
    });
    await enterCommand(page, "Try again");

    const turnError = await page.waitForFunction(() => {
      const sidebar = document.querySelector("pi-sidebar");
      const agent = sidebar && "agent" in sidebar
        ? sidebar.agent as { state: { messages: { errorMessage?: string }[] } }
        : null;
      const failed = agent?.state.messages.some((message) => message.errorMessage === "Bad request: unsupported parameter");
      return failed && sidebar && "turnError" in sidebar && typeof sidebar.turnError === "string" ? sidebar.turnError : null;
    });
    assert.equal(await turnError.jsonValue(), "Pi couldn’t finish this response. Try again or choose another model.");
  } finally {
    await opened.finish();
  }
});
