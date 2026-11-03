import { authDeferred } from "../fixtures/auth-deferred.ts";
import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import type { Page, Route } from "playwright";

import { openBrowserPage, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
const previousNodeEnv = process.env.NODE_ENV;
before(async () => {
  process.env.NODE_ENV = "production";
  env = await startTaskpaneServer();
});
after(async () => {
  await env.close();
  if (previousNodeEnv === undefined) delete process.env.NODE_ENV;
  else process.env.NODE_ENV = previousNodeEnv;
});

async function withProviders(run: (page: Page) => Promise<void>, active = false): Promise<void> {
  const opened = await openBrowserPage(env, {
    path: "auth-fixture",
    prepareContext: async (context) => {
      await context.route("**/auth-fixture", (route) => route.fulfill({
        contentType: "text/html",
        body: '<html><head><link rel="stylesheet" href="/src/ui/theme.css"></head><body></body></html>',
      }));
      await context.route("https://localhost:3003/**", (route) => route.fulfill({ status: 200, body: "ok" }));
    },
  });
  try {
    await opened.page.evaluate(async (active) => {
      const storagePath = "/src/storage/init-app-storage.ts";
      const rowPath = "/src/ui/provider-login.ts";
      const { initAppStorage } = await import(/* @vite-ignore */ storagePath) as typeof import("../../src/storage/init-app-storage.ts");
      const { buildProviderRow } = await import(/* @vite-ignore */ rowPath) as typeof import("../../src/ui/provider-login.ts");
      const { storage } = initAppStorage("provider-login-browser");
      await storage.settings.set("proxy.enabled", true);
      await storage.settings.set("proxy.url", "https://localhost:3003");
      if (active) await storage.providerKeys.set("anthropic", "existing-fixture");
      const lifetime = new AbortController();
      const dispose = document.createElement("button");
      dispose.textContent = "Dispose providers";
      dispose.addEventListener("click", () => { lifetime.abort(); document.querySelectorAll(".pi-login-row").forEach((row) => row.remove()); });
      document.body.append(dispose);
      const expandedRef = { current: null };
      for (const id of ["anthropic", "openai"]) {
        const row = buildProviderRow({ id, brand: id, ...(id === "anthropic" ? { oauth: id } : {}) }, {
          signal: lifetime.signal, expandedRef, isActive: active && id === "anthropic",
          onConnected: () => { document.body.dataset.connected = id; },
          onDisconnected: () => { document.body.dataset.disconnected = id; },
        });
        row.dataset.provider = id;
        document.body.append(row);
      }
    }, active);
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

async function readCredentials(page: Page): Promise<{ key: string | null; grant: boolean }> {
  return page.evaluate(async () => {
    const path = "/src/storage/local/app-storage.ts";
    const { getAppStorage } = await import(/* @vite-ignore */ path) as typeof import("../../src/storage/local/app-storage.ts");
    const storage = getAppStorage();
    return { key: await storage.providerKeys.get("anthropic"), grant: Boolean(await storage.settings.get("oauth.anthropic")) };
  });
}

async function startKeySave(page: Page): Promise<Route> {
  const pending = authDeferred<Route>();
  await page.route("https://api.anthropic.com/v1/models", (route) => { pending.resolve(route); });
  const row = page.locator('[data-provider="anthropic"]');
  await row.locator(".pi-login-trigger").click();
  await row.locator(".pi-login-key").fill("fixture-key");
  await row.locator(".pi-login-save").click();
  return pending.promise;
}

async function startOAuthExchange(page: Page): Promise<Route> {
  const pending = authDeferred<Route>();
  await page.route("https://platform.claude.com/v1/oauth/token", (route) => { pending.resolve(route); });
  const row = page.locator('[data-provider="anthropic"]');
  await row.locator(".pi-login-trigger").click();
  await row.locator(".pi-login-oauth").click();
  const prompt = page.locator("#pi-prompt-overlay");
  await prompt.locator(".pi-prompt-input").fill("fixture-code");
  await prompt.locator(".pi-prompt-ok").click();
  return pending.promise;
}

async function finishToken(route: Route): Promise<void> {
  await route.fulfill({ json: { access_token: "fixture-access", refresh_token: "fixture-refresh", expires_in: 3600 } });
}

void test("an OAuth operation exclusively owns its row controls", async () => {
  await withProviders(async (page) => {
    const row = page.locator('[data-provider="anthropic"]');
    await row.locator(".pi-login-trigger").click();
    await row.locator(".pi-login-oauth").click();
    await page.locator(".pi-prompt-input").waitFor();
    assert.equal(await row.locator(".pi-login-oauth").isDisabled(), true, "OAuth button must reject another login");
    assert.equal(await row.locator(".pi-login-save").isDisabled(), true, "key save must not overlap OAuth");
    await page.locator(".pi-prompt-cancel").click();
    await page.waitForFunction(() => !document.querySelector<HTMLButtonElement>(".pi-login-oauth")?.disabled);
    assert.deepEqual(await readCredentials(page), { key: null, grant: false });
  });
});

void test("disconnect cancels a pending key validation instead of reconnecting later", async () => {
  await withProviders(async (page) => {
    const validation = await startKeySave(page);
    await page.locator(".pi-login-disconnect").first().click();
    await page.waitForFunction(() => document.body.dataset.disconnected === "anthropic");
    await validation.fulfill({ status: 200, body: "ok" });
    await page.waitForFunction(() => !document.querySelector<HTMLButtonElement>(".pi-login-save")?.disabled);
    assert.deepEqual(await readCredentials(page), { key: null, grant: false }, "late key validation must not reconnect");
    assert.equal(await page.locator("body").getAttribute("data-connected"), null);
  }, true);
});

void test("save completion preserves another row's expansion ownership", async () => {
  await withProviders(async (page) => {
    const validation = await startKeySave(page);
    const other = page.locator('[data-provider="openai"]');
    await other.locator(".pi-login-trigger").click();
    await validation.fulfill({ status: 200, body: "ok" });
    await page.waitForFunction(() => document.body.dataset.connected === "anthropic");
    await other.locator(".pi-login-trigger").click();
    assert.equal(await other.locator(".pi-login-detail").isVisible(), false, "expanded row must still own its toggle");
  });
});

void test("closing the API key dialog cancels validation before credentials are persisted", async () => {
  await withProviders(async (page) => {
    await page.getByRole("button", { name: "Dispose providers" }).click();
    await page.evaluate(async () => {
      const path = "/src/ui/api-key-dialog.ts";
      const { promptForProviderConnection } = await import(/* @vite-ignore */ path) as typeof import("../../src/ui/api-key-dialog.ts");
      void promptForProviderConnection("anthropic").then((connected) => { document.body.dataset.dialogResult = String(connected); });
    });
    const pending = authDeferred<Route>();
    await page.route("https://api.anthropic.com/v1/models", (route) => { pending.resolve(route); });
    await page.locator(".pi-login-key").fill("fixture-key");
    await page.locator(".pi-login-save").click();
    const validation = await pending.promise;
    await page.keyboard.press("Escape");
    await page.waitForFunction(() => document.body.dataset.dialogResult === "false");
    await validation.fulfill({ status: 200, body: "ok" });
    await page.waitForFunction(() => !document.querySelector(".pi-api-key-card"));
    await readCredentials(page);
    assert.deepEqual(await readCredentials(page), { key: null, grant: false }, "disposed dialog must not save a key");
  });
});

void test("disconnect cancels an OAuth token exchange and prevents late credential writes", async () => {
  await withProviders(async (page) => {
    const loginLabel = await page.locator(".pi-login-oauth").textContent();
    const exchange = await startOAuthExchange(page);
    await page.locator(".pi-login-disconnect").first().click();
    await page.waitForFunction(() => document.body.dataset.disconnected === "anthropic");
    await finishToken(exchange);
    await page.waitForFunction((label) => document.querySelector(".pi-login-oauth")?.textContent === label, loginLabel);
    assert.deepEqual(await readCredentials(page), { key: null, grant: false }, "cancelled exchange must not reconnect");
  }, true);
});

void test("logout shares the refresh lock so an in-flight refresh cannot resurrect credentials", async () => {
  await withProviders(async (page) => {
    const pending = authDeferred<Route>();
    await page.route("https://platform.claude.com/v1/oauth/token", (route) => { pending.resolve(route); });
    await page.evaluate(async () => {
      const storagePath = "/src/storage/local/app-storage.ts";
      const refreshPath = "/src/auth/oauth-refresh.ts";
      const providersPath = "/src/auth/oauth-provider-registry.ts";
      const { getAppStorage } = await import(/* @vite-ignore */ storagePath) as typeof import("../../src/storage/local/app-storage.ts");
      const { createOAuthRefresher } = await import(/* @vite-ignore */ refreshPath) as typeof import("../../src/auth/oauth-refresh.ts");
      const { getOAuthProvider } = await import(/* @vite-ignore */ providersPath) as typeof import("../../src/auth/oauth-provider-registry.ts");
      const storage = getAppStorage();
      await storage.settings.set("oauth.anthropic", { access: "old-access", refresh: "old-refresh", expires: 1 });
      const refresher = createOAuthRefresher({ settings: storage.settings, providerKeys: storage.providerKeys, getOAuthProvider });
      void refresher.ensureFresh("anthropic").then(() => { document.body.dataset.refresh = "finished"; });
    });
    const refresh = await pending.promise;
    await page.locator(".pi-login-trigger").first().click();
    await page.locator(".pi-login-disconnect").first().click();
    await readCredentials(page);
    await readCredentials(page);
    await finishToken(refresh);
    await page.waitForFunction(() => document.body.dataset.refresh === "finished" && document.body.dataset.disconnected === "anthropic");
    assert.deepEqual(await readCredentials(page), { key: null, grant: false }, "refresh must finish before logout deletes credentials");
  }, true);
});

for (const retryTwice of [false, true]) {
  void test(`proxy retry ${retryTwice ? "is single-flight" : "cannot enable the proxy after cancellation"}`, async () => {
    await withProviders(async (page) => {
      await page.evaluate(async () => {
        const path = "/src/storage/local/app-storage.ts";
        const { getAppStorage } = await import(/* @vite-ignore */ path) as typeof import("../../src/storage/local/app-storage.ts");
        await getAppStorage().settings.set("proxy.enabled", false);
      });
      const pending = authDeferred<Route>();
      let probes = 0;
      await page.route("https://localhost:3003/?url=**", (route) => { probes += 1; pending.resolve(route); });
      await page.locator(".pi-login-trigger").first().click();
      await page.locator(".pi-login-oauth").click();
      const gate = page.locator("#pi-proxy-gate-overlay");
      await gate.locator(".pi-prompt-ok").click();
      const probe = await pending.promise;
      if (retryTwice) {
        await gate.locator(".pi-prompt-ok").dispatchEvent("click");
        await page.evaluate(() => Promise.resolve());
        assert.equal(probes, 1, "retry must not launch overlapping probes");
      }
      await gate.locator(".pi-prompt-cancel").click();
      await probe.fulfill({ status: 200, body: "ok" });
      await page.waitForFunction(() => !document.querySelector<HTMLButtonElement>(".pi-login-oauth")?.disabled);
      const enabled = await page.evaluate(async () => {
        const path = "/src/storage/local/app-storage.ts";
        const { getAppStorage } = await import(/* @vite-ignore */ path) as typeof import("../../src/storage/local/app-storage.ts");
        return getAppStorage().settings.get("proxy.enabled");
      });
      assert.equal(enabled, false, "a cancelled probe must not enable the proxy");
    });
  });
}
