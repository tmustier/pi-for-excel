import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { chromium, type Browser } from "playwright";
import { createServer, type ViteDevServer } from "vite";

let browser: Browser;
let server: ViteDevServer;
let baseUrl: string;

before(async () => {
  process.env.VITE_PI_BACKGROUND_VERIFY_URL = "https://localhost:3157";
  process.env.VITE_PI_BACKGROUND_VERIFY_TOKEN = "extensions-browser-token";

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

void test("an inline extension installed through the runtime appears enabled in Plugins", async () => {
  const context = await browser.newContext();
  const page = await context.newPage();
  const pageErrors: Error[] = [];
  let commandSent = false;
  let resolveInstalled: (() => void) | undefined;
  const installed = new Promise<void>((resolve) => {
    resolveInstalled = resolve;
  });

  page.on("pageerror", (error) => {
    if (!isExpectedOfficeUnavailableError(error)) pageErrors.push(error);
  });

  await context.route("**/*", async (route) => {
    const url = new URL(route.request().url());
    if (url.hostname === "localhost" && url.port === "3157") {
      if (url.pathname === "/client/register") {
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ clientId: "extension-hub-browser" }) });
      } else if (url.pathname === "/client/poll" && !commandSent) {
        commandSent = true;
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify({
            id: "install-inline-extension",
            type: "extensionInstallCode",
            payload: {
              name: "Browser audit extension",
              code: "export function activate() {}",
            },
          }),
        });
      } else if (url.pathname === "/client/result") {
        resolveInstalled?.();
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true }) });
      } else {
        await new Promise((resolve) => setTimeout(resolve, 250));
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ type: "noop" }) });
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

    await Promise.race([
      installed,
      new Promise<void>((_resolve, reject) => {
        setTimeout(() => reject(new Error("inline extension installation did not complete")), 10_000);
      }),
    ]);

    await input.fill("/plugins");
    await input.press("Enter");

    const plugins = page.locator("#pi-settings-overlay");
    await plugins.getByRole("heading", { name: "Plugins" }).waitFor({ state: "visible", timeout: 5_000 });
    const extensionCard = plugins.locator(".pi-item-card", { hasText: "Browser audit extension" });
    await extensionCard.waitFor({ state: "visible" });
    assert.match((await extensionCard.textContent()) ?? "", /inline code \(29 chars, 1 lines\) · sandbox iframe/u);
    await assert.doesNotReject(() => extensionCard.getByRole("checkbox").waitFor({ state: "attached" }));
    assert.equal(await extensionCard.getByRole("checkbox").isChecked(), true);
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});
