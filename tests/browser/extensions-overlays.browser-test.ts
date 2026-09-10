import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { chromium, type Browser, type BrowserContext, type Page } from "playwright";
import { createServer, type ViteDevServer } from "vite";

let browser: Browser;
let server: ViteDevServer;
let baseUrl: string;

before(async () => {
  process.env.VITE_PI_BACKGROUND_VERIFY_URL = "https://localhost:3157";
  process.env.VITE_PI_BACKGROUND_VERIFY_TOKEN = "extension-overlays-browser-token";

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
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ clientId: "extension-overlay-test" }) });
      } else if (url.pathname === "/client/poll" && !commandSent) {
        commandSent = true;
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify({
            id: "install-overlay-extension",
            type: "extensionInstallCode",
            payload: { name: "Overlay browser extension", code },
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
  await welcomeOverlay.click({ position: { x: 2, y: 2 }, force: true });
  await welcomeOverlay.waitFor({ state: "detached" });

  return { context, page, pageErrors };
}

void test("an installed extension can replace and dismiss its visible overlay", async () => {
  const { context, page, pageErrors } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const first = document.createElement("section");
      first.textContent = "First extension overlay";
      api.overlay.show(first);

      setTimeout(() => {
        const second = document.createElement("section");
        second.textContent = "Replacement extension overlay";
        api.overlay.show(second);
      }, 1000);

      setTimeout(() => api.overlay.dismiss(), 2000);
    }
  `);

  try {
    const overlay = page.locator("#pi-ext-overlay");
    await overlay.getByText("First extension overlay").waitFor({ state: "visible", timeout: 10_000 });

    await overlay.getByText("Replacement extension overlay").waitFor({ state: "visible", timeout: 5_000 });
    assert.equal(await overlay.getByText("First extension overlay").count(), 0);

    await overlay.waitFor({ state: "detached", timeout: 5_000 });
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});
