import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { chromium, type Browser, type BrowserContext, type Page } from "playwright";
import { createServer, type ViteDevServer } from "vite";

let browser: Browser;
let server: ViteDevServer;
let baseUrl: string;

before(async () => {
  server = await createServer({
    configFile: false,
    root: process.cwd(),
    server: {
      host: "127.0.0.1",
      port: 0,
      strictPort: false,
    },
  });
  await server.listen();

  const localUrl = server.resolvedUrls?.local[0];
  if (!localUrl) {
    throw new Error("Vite did not expose a local URL");
  }
  baseUrl = localUrl;
  browser = await chromium.launch({ headless: true });
});

after(async () => {
  await browser?.close();
  await server?.close();
});

function isLoopbackUrl(rawUrl: string): boolean {
  const url = new URL(rawUrl);
  return url.hostname === "127.0.0.1" || url.hostname === "localhost" || url.hostname === "[::1]";
}

function isExpectedOfficeUnavailableError(error: Error): boolean {
  return /Office(?:\.js)? (?:is |was )?(?:not ready|not available|unavailable)/i.test(error.message);
}

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  const context: BrowserContext = await browser.newContext();
  const page = await context.newPage();
  const pageErrors: Error[] = [];

  page.on("pageerror", (error) => {
    if (!isExpectedOfficeUnavailableError(error)) {
      pageErrors.push(error);
    }
  });

  await context.route("**/*", async (route) => {
    if (isLoopbackUrl(route.request().url())) {
      await route.continue();
      return;
    }
    await route.abort("blockedbyclient");
  });

  try {
    await page.goto(`${baseUrl}src/taskpane.html`);
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });

    const welcomeOverlay = page.locator("#pi-welcome-login-overlay");
    await welcomeOverlay.waitFor({ state: "visible", timeout: 10_000 });
    await welcomeOverlay.click({ position: { x: 2, y: 2 } });
    await welcomeOverlay.waitFor({ state: "detached" });

    await run(page);
    assert.deepEqual(pageErrors.map((error) => error.message), [], "unexpected uncaught page errors");
  } finally {
    await context.close();
  }
}

async function enterCommand(page: Page, command: string): Promise<void> {
  const input = page.locator("pi-input textarea");
  await input.fill(command);
  await input.press("Enter");
}

async function openUtilitiesMenu(page: Page): Promise<void> {
  await page.getByRole("button", { name: "Settings and tools" }).click();
}

void test("/files opens the Files view", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/files");
    await page.getByRole("heading", { name: "Files" }).waitFor({ state: "visible" });
  });
});

void test("/plugins opens the extensions hub", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/plugins");
    await page.getByRole("heading", { name: "Connections" }).waitFor({ state: "visible", timeout: 5_000 });
  });
});

async function verifyFailedCommandToast(failureKind: "rejects" | "throws"): Promise<void> {
  await withTaskpane(async (page) => {
    const sentinel = `browser-command-${failureKind}-sentinel`;
    const executeSource = failureKind === "rejects"
      ? `() => Promise.reject(new Error(${JSON.stringify(sentinel)}))`
      : `() => { throw new Error(${JSON.stringify(sentinel)}); }`;
    await page.evaluate(`
      (async () => {
        const { commandRegistry } = await import("/src/commands/types.ts");
        commandRegistry.register({
          name: "browser-${failureKind}",
          description: "Browser failure contract",
          source: "extension",
          execute: ${executeSource},
        });
      })()
    `);

    await enterCommand(page, `/browser-${failureKind}`);
    const toast = page.locator("#pi-toast.visible .pi-toast__message");
    await toast.waitFor({ state: "visible", timeout: 5_000 });
    assert.equal(await toast.textContent(), "Could not run that command.");
    assert.equal((await page.locator("body").textContent())?.includes(sentinel), false);
  });
}

void test("a command that rejects shows only the generic failure toast", async () => {
  await verifyFailedCommandToast("rejects");
});

void test("a command that throws shows only the generic failure toast", async () => {
  await verifyFailedCommandToast("throws");
});

void test("sidebar Settings opens the unified settings overlay", async () => {
  await withTaskpane(async (page) => {
    await openUtilitiesMenu(page);
    await page.getByRole("menuitem", { name: "Settings" }).click();
    await page.getByRole("heading", { name: "Settings" }).waitFor({ state: "visible", timeout: 5_000 });
  });
});

void test("sidebar Files and Extensions buttons open their views", async () => {
  await withTaskpane(async (page) => {
    await openUtilitiesMenu(page);
    await page.getByRole("menuitem", { name: "Files" }).click();
    await page.getByRole("heading", { name: "Files" }).waitFor({ state: "visible", timeout: 5_000 });
    await page.getByRole("button", { name: "Close Files" }).click();

    await openUtilitiesMenu(page);
    await page.getByRole("menuitem", { name: "Extensions" }).click();
    await page.getByRole("heading", { name: "Connections" }).waitFor({ state: "visible", timeout: 5_000 });
  });
});

void test("proxy state changes update the proxy banner", async () => {
  await withTaskpane(async (page) => {
    const banner = page.locator(".pi-proxy-banner");
    await banner.waitFor({ state: "attached" });
    await page.evaluate(() => {
      document.dispatchEvent(new CustomEvent("pi:proxy-state-changed", {
        detail: { state: "not-detected" },
      }));
    });
    await banner.waitFor({ state: "visible" });
    assert.match(await banner.innerText(), /Proxy not running/);

    await page.evaluate(() => {
      document.dispatchEvent(new CustomEvent("pi:proxy-state-changed", {
        detail: { state: "detected" },
      }));
    });
    await banner.waitFor({ state: "hidden" });
  });
});

void test("status bar exposes model, thinking, context, and mode controls only", async () => {
  await withTaskpane(async (page) => {
    for (const selector of [".pi-status-model", ".pi-status-thinking", ".pi-status-ctx", ".pi-status-mode"]) {
      await page.locator(selector).waitFor({ state: "visible" });
    }
    assert.equal(await page.locator(".pi-status-rules").count(), 0);
    assert.equal(await page.locator(".pi-status-proxy").count(), 0);
  });
});
