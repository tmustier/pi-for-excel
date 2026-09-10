import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { chromium, type Browser, type BrowserContext, type Page } from "playwright";
import { createServer, type ViteDevServer } from "vite";

let browser: Browser;
let server: ViteDevServer;
let baseUrl: string;

before(async () => {
  process.env.VITE_PI_BACKGROUND_VERIFY_URL = "https://localhost:3157";
  process.env.VITE_PI_BACKGROUND_VERIFY_TOKEN = "browser-bridge-token";

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

async function withTaskpane(run: (page: Page) => Promise<void>, viewportHeight = 720): Promise<void> {
  const context: BrowserContext = await browser.newContext({ viewport: { width: 420, height: viewportHeight } });
  const page = await context.newPage();
  const pageErrors: Error[] = [];

  page.on("pageerror", (error) => {
    if (!isExpectedOfficeUnavailableError(error)) pageErrors.push(error);
  });

  await context.route("**/*", async (route) => {
    const url = new URL(route.request().url());
    if (url.hostname === "localhost" && url.port === "3157") {
      if (url.pathname === "/client/register") {
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ clientId: "composer-browser-client" }) });
      } else if (url.pathname === "/client/poll") {
        await new Promise((resolve) => setTimeout(resolve, 1_000));
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ type: "noop" }) });
      } else {
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true }) });
      }
      return;
    }

    if (url.hostname === "127.0.0.1" || url.hostname === "localhost" || url.hostname === "[::1]") {
      await route.continue();
      return;
    }
    await route.abort("blockedbyclient");
  });

  try {
    await page.goto(`${baseUrl}src/taskpane.html`);
    const input = page.locator("pi-input textarea");
    await input.waitFor({ state: "visible", timeout: 20_000 });

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

async function enterInput(page: Page, text: string): Promise<void> {
  const input = page.locator("pi-input textarea");
  await input.fill(text);
  await input.press("Enter");
}

void test("Alt+Up recalls queued actions into the composer in order", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(`
      (async () => {
        const { commandRegistry } = await import("/src/commands/types.ts");
        commandRegistry.register({
          name: "compact",
          description: "Keep the action queue busy",
          source: "builtin",
          execute: () => new Promise(() => {}),
        });
      })()
    `);

    await enterInput(page, "/compact");
    await enterInput(page, "first queued action");
    await page.locator("#pi-queue-display").getByText("first queued action", { exact: true }).waitFor({ state: "visible" });
    await enterInput(page, "second queued action");
    await page.locator("#pi-queue-display").getByText("second queued action", { exact: true }).waitFor({ state: "visible" });

    const input = page.locator("pi-input textarea");
    await input.press("Alt+ArrowUp");

    assert.deepEqual(
      {
        composer: await input.inputValue(),
        queuedItems: await page.locator("#pi-queue-display .pi-queue__item").count(),
      },
      {
        composer: "first queued action\n\nsecond queued action",
        queuedItems: 0,
      },
    );
  });
});

void test("the composer trims a prompt and does not submit whitespace", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await page.locator("pi-input").evaluate((composer) => {
      composer.addEventListener("pi-send", (event) => {
        // The component owns this event contract; the assertion only exposes its browser payload for observation.
        const sendEvent = event as CustomEvent<{ text?: string }>;
        if (typeof sendEvent.detail?.text === "string") {
          document.body.dataset.lastComposerSend = sendEvent.detail.text;
        }
      });
    });
    await input.fill("   \n\t  ");
    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-last-composer-send"), null);

    await input.fill("  build a TSLA P&L  ");
    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-last-composer-send"), "build a TSLA P&L");
  });
});

void test("Shift+Enter edits while Enter submits a real prompt", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await page.locator("pi-input").evaluate((composer) => {
      composer.addEventListener("pi-send", () => { document.body.dataset.composerSent = "yes"; });
    });
    await input.fill("write A1");
    await input.press("Shift+Enter");
    assert.equal(await page.locator("body").getAttribute("data-composer-sent"), null);
    assert.equal(await input.inputValue(), "write A1\n");

    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-composer-sent"), "yes");
  });
});

void test("composer auto-grow honors its CSS cap", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.fill(Array.from({ length: 50 }, () => "row").join("\n"));
    assert.equal(await input.evaluate((element) => element.style.height), "288px");
  }, 720);
});

void test("composer auto-grow uses the smaller short-pane cap", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.evaluate((element) => { element.style.maxHeight = "none"; });
    await input.fill(Array.from({ length: 50 }, () => "row").join("\n"));
    assert.equal(await input.evaluate((element) => element.style.height), "140px");
  }, 500);
});
