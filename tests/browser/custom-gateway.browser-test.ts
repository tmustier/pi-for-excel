import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { chromium, type Browser, type Page } from "playwright";
import { createServer, type ViteDevServer } from "vite";

let browser: Browser;
let server: ViteDevServer;
let baseUrl: string;

before(async () => {
  process.env.VITE_PI_BACKGROUND_VERIFY_URL = "https://localhost:3157";
  process.env.VITE_PI_BACKGROUND_VERIFY_TOKEN = "browser-bridge-token";
  server = await createServer({ configFile: false, root: process.cwd(), server: { host: "127.0.0.1", port: 0 } });
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

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  const context = await browser.newContext();
  const page = await context.newPage();
  await context.route("**/*", async (route) => {
    const url = new URL(route.request().url());
    if (url.hostname === "localhost" && url.port === "3157") {
      const body = url.pathname === "/client/register" ? { clientId: "gateway-browser" } : { type: "noop" };
      await route.fulfill({ contentType: "application/json", body: JSON.stringify(body) });
    } else if (["127.0.0.1", "localhost", "[::1]"].includes(url.hostname)) {
      await route.continue();
    } else {
      await route.abort("blockedbyclient");
    }
  });

  try {
    await page.goto(`${baseUrl}src/taskpane.html`);
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    const welcome = page.locator("#pi-welcome-login-overlay");
    await welcome.waitFor({ state: "visible", timeout: 10_000 });
    await welcome.click({ position: { x: 2, y: 2 } });
    await welcome.waitFor({ state: "detached" });
    await run(page);
  } finally {
    await context.close();
  }
}

void test("custom gateway deletion persists only after overlay confirmation", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.fill("/settings");
    await input.press("Enter");
    const settings = page.locator("#pi-settings-overlay");
    await settings.getByRole("button", { name: "Custom gateway" }).click();

    await page.getByPlaceholder("Gateway name (optional)").fill("Browser gateway");
    await page.getByPlaceholder("https://your-gateway.example.com/v1").fill("https://gateway.example.com/v1");
    await page.getByPlaceholder("model-id").fill("browser-model");
    await page.getByRole("button", { name: "Save gateway" }).click();
    const gateway = settings.getByText("Browser gateway", { exact: true });
    await gateway.waitFor({ state: "visible" });

    await settings.getByRole("button", { name: "Delete" }).click();
    const confirmation = page.locator("#pi-confirm-dialog-overlay");
    await confirmation.locator(".pi-overlay-btn--ghost").click();
    await confirmation.waitFor({ state: "detached" });
    await page.waitForTimeout(250);
    assert.equal(await gateway.isVisible(), true);

    await settings.getByRole("button", { name: "Delete" }).click();
    await confirmation.getByRole("button", { name: "Delete" }).click();
    await gateway.waitFor({ state: "detached" });
    assert.equal(await settings.getByText("No custom gateways configured yet.", { exact: true }).isVisible(), true);
  });
});
