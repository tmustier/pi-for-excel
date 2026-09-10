import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { chromium, type Browser, type BrowserContext, type Page } from "playwright";
import { createServer, type ViteDevServer } from "vite";

let browser: Browser;
let server: ViteDevServer;
let baseUrl: string;

before(async () => {
  process.env.VITE_PI_BACKGROUND_VERIFY_URL = "https://localhost:3157";
  process.env.VITE_PI_BACKGROUND_VERIFY_TOKEN = "extension-widgets-browser-token";

  server = await createServer({
    configFile: false,
    root: process.cwd(),
    server: { host: "127.0.0.1", port: 0, strictPort: false },
  });
  await server.listen();

  const localUrl = server.resolvedUrls?.local[0];
  if (!localUrl) throw new Error("Vite did not expose a local URL");
  baseUrl = localUrl;
  browser = await chromium.launch({ headless: true });
});

after(async () => {
  await browser?.close();
  await server?.close();
});

function isExpectedOfficeUnavailableError(error: Error): boolean {
  return /Office(?:\.js)? (?:is |was )?(?:not ready|not available|unavailable)/i.test(error.message);
}

async function openTaskpaneWithExtension(code: string): Promise<{
  context: BrowserContext;
  page: Page;
  pageErrors: Error[];
}> {
  const context = await browser.newContext();
  await context.addInitScript(() => {
    if (window === window.top) {
      localStorage.setItem("pi.experimental.extensionWidgetV2", "1");
    }
  });
  const page = await context.newPage();
  const pageErrors: Error[] = [];
  let commandSent = false;

  page.on("pageerror", (error) => {
    if (!isExpectedOfficeUnavailableError(error)) pageErrors.push(error);
  });

  await context.route("**/*", async (route) => {
    const url = new URL(route.request().url());
    if (url.hostname === "localhost" && url.port === "3157") {
      if (url.pathname === "/client/register") {
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ clientId: "extension-widget-test" }) });
      } else if (url.pathname === "/client/poll" && !commandSent) {
        commandSent = true;
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify({
            id: "install-widget-extension",
            type: "extensionInstallCode",
            payload: { name: "Widget browser extension", code },
          }),
        });
      } else {
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true, type: "noop" }) });
      }
      return;
    }

    if (url.hostname === "127.0.0.1" || url.hostname === "localhost" || url.hostname === "[::1]") {
      await route.continue();
      return;
    }
    await route.abort("blockedbyclient");
  });

  await page.goto(`${baseUrl}src/taskpane.html`);
  await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
  const welcomeOverlay = page.locator("#pi-welcome-login-overlay");
  await welcomeOverlay.waitFor({ state: "visible", timeout: 10_000 });
  await welcomeOverlay.click({ position: { x: 2, y: 2 } });
  await welcomeOverlay.waitFor({ state: "detached" });

  return { context, page, pageErrors };
}

void test("an installed extension's widget bounds are clamped and remain valid", async () => {
  const { context, page, pageErrors } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const el = document.createElement("div");
      el.textContent = "Bounded widget content";
      api.widget.upsert({
        id: "bounded",
        el,
        title: "Bounded widget",
        minHeightPx: 700,
        maxHeightPx: 12
      });
    }
  `);

  try {
    const body = page.locator("[data-widget-id='bounded'] .pi-ext-widget-body");
    await body.waitFor({ state: "visible", timeout: 10_000 });
    assert.equal(await body.evaluate((element) => element.style.minHeight), "640px");
    assert.equal(await body.evaluate((element) => element.style.maxHeight), "640px");
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});

void test("an installed extension preserves widget bounds when an upsert omits them", async () => {
  const { context, page, pageErrors } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const first = document.createElement("div");
      first.textContent = "First bounded content";
      api.widget.upsert({ id: "preserved", el: first, minHeightPx: 180, maxHeightPx: 420 });

      const replacement = document.createElement("div");
      replacement.textContent = "Replacement bounded content";
      api.widget.upsert({ id: "preserved", el: replacement });
    }
  `);

  try {
    const body = page.locator("[data-widget-id='preserved'] .pi-ext-widget-body");
    await body.waitFor({ state: "visible", timeout: 10_000 });
    assert.deepEqual(
      await body.evaluate((element) => ({ min: element.style.minHeight, max: element.style.maxHeight })),
      { min: "180px", max: "420px" },
    );
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});

void test("an installed extension can clear existing widget bounds", async () => {
  const { context, page, pageErrors } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const first = document.createElement("div");
      first.textContent = "Initially bounded content";
      api.widget.upsert({ id: "cleared", el: first, minHeightPx: 180, maxHeightPx: 420 });

      const replacement = document.createElement("div");
      replacement.textContent = "Unbounded replacement content";
      api.widget.upsert({ id: "cleared", el: replacement, minHeightPx: null, maxHeightPx: null });
    }
  `);

  try {
    const body = page.locator("[data-widget-id='cleared'] .pi-ext-widget-body");
    await body.waitFor({ state: "visible", timeout: 10_000 });
    assert.deepEqual(
      await body.evaluate((element) => ({ min: element.style.minHeight, max: element.style.maxHeight })),
      { min: "", max: "" },
    );
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});

void test("an installed extension cannot leave a non-collapsible widget collapsed", async () => {
  const { context, page, pageErrors } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const el = document.createElement("div");
      el.textContent = "Always visible widget content";
      api.widget.upsert({
        id: "always-visible",
        el,
        title: "Always visible widget",
        collapsible: false,
        collapsed: true
      });
    }
  `);

  try {
    const card = page.locator("[data-widget-id='always-visible']");
    const body = card.locator(".pi-ext-widget-body");
    await body.waitFor({ state: "visible", timeout: 10_000 });
    assert.equal(await card.getByRole("button").count(), 0);
    assert.equal(await body.isVisible(), true);
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});

void test("an installed extension preserves collapsed state when an upsert omits it", async () => {
  const { context, page, pageErrors } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const first = document.createElement("div");
      first.textContent = "Initially collapsed content";
      api.widget.upsert({ id: "collapsed", el: first, title: "Collapsed widget", collapsible: true, collapsed: true });

      const replacement = document.createElement("div");
      replacement.textContent = "Replacement collapsed content";
      api.widget.upsert({ id: "collapsed", el: replacement });
    }
  `);

  try {
    const card = page.locator("[data-widget-id='collapsed']");
    await card.waitFor({ state: "visible", timeout: 10_000 });
    assert.equal(await card.locator(".pi-ext-widget-body").isVisible(), false);
    assert.equal(await card.getByRole("button").getAttribute("aria-expanded"), "false");
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});
