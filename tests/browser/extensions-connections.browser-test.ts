import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { chromium, type Browser, type BrowserContext, type Page } from "playwright";
import { createServer, type ViteDevServer } from "vite";

let browser: Browser;
let server: ViteDevServer;
let baseUrl: string;

before(async () => {
  process.env.VITE_PI_BACKGROUND_VERIFY_URL = "https://localhost:3157";
  process.env.VITE_PI_BACKGROUND_VERIFY_TOKEN = "extension-connections-browser-token";
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

function readInstalledExtensionId(raw: DynamicValue): string | null {
  if (typeof raw !== "object" || raw === null || Array.isArray(raw)) return null;
  const envelope = raw as DynamicObject;
  if (typeof envelope.result !== "object" || envelope.result === null || Array.isArray(envelope.result)) return null;
  const result = envelope.result as DynamicObject;
  return typeof result.extensionId === "string" ? result.extensionId : null;
}

async function openConnectionsWithExtension(code?: string, grantConnections = false): Promise<{
  context: BrowserContext;
  page: Page;
  plugins: ReturnType<Page["locator"]>;
  pageErrors: Error[];
}> {
  const context = await browser.newContext();
  const page = await context.newPage();
  const pageErrors: Error[] = [];
  let commandSent = false;
  let capabilitySent = false;
  let installedExtensionId: string | null = null;
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
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ clientId: "extension-connections-test" }) });
      } else if (url.pathname === "/client/poll" && !commandSent) {
        commandSent = true;
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify(code
            ? {
              id: "install-connection-extension",
              type: "extensionInstallCode",
              payload: { name: "Connection browser extension", code },
            }
            : {
              id: "uninstall-builtin-extension",
              type: "extensionUninstall",
              payload: { extensionId: "builtin.snake" },
            }),
        });
      } else if (grantConnections && installedExtensionId && url.pathname === "/client/poll" && !capabilitySent) {
        capabilitySent = true;
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify({
            id: "grant-connection-capability",
            type: "extensionSetCapability",
            payload: {
              extensionId: installedExtensionId,
              capability: "connections.readwrite",
              allowed: true,
            },
          }),
        });
      } else if (url.pathname === "/client/result") {
        const raw: DynamicValue = JSON.parse(route.request().postData() ?? "{}");
        installedExtensionId = installedExtensionId ?? readInstalledExtensionId(raw);
        if (!grantConnections || capabilitySent) resolveInstalled?.();
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true }) });
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
  const input = page.locator("pi-input textarea");
  await input.waitFor({ state: "visible", timeout: 20_000 });
  const welcomeOverlay = page.locator("#pi-welcome-login-overlay");
  await welcomeOverlay.waitFor({ state: "visible", timeout: 10_000 });
  await welcomeOverlay.click({ position: { x: 2, y: 2 } });
  await welcomeOverlay.waitFor({ state: "detached" });

  await Promise.race([
    installed,
    new Promise<void>((_resolve, reject) => {
      setTimeout(() => reject(new Error("extension setup did not complete")), 10_000);
    }),
  ]);

  await input.fill("/plugins");
  await input.press("Enter");
  const plugins = page.locator("#pi-settings-overlay");
  await plugins.getByRole("heading", { name: "Plugins" }).waitFor({ state: "visible", timeout: 5_000 });
  await plugins.getByRole("button", { name: "Back" }).click();
  await plugins.getByRole("button", { name: /^Connections/u }).click();
  await plugins.getByRole("heading", { name: "Connections" }).waitFor({ state: "visible" });
  return { context, page, plugins, pageErrors };
}

void test("Connections hides extension connections when no extensions are installed", async () => {
  const { context, plugins, pageErrors } = await openConnectionsWithExtension();
  try {
    await plugins.locator(".pi-section-header__label", { hasText: "MCP servers" }).waitFor({ state: "visible" });
    assert.equal(await plugins.locator(".pi-section-header__label", { hasText: "Extension connections" }).count(), 0);
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});

void test("Connections explains when installed extensions register no connections", async () => {
  const { context, plugins, pageErrors } = await openConnectionsWithExtension("export function activate() {}");
  try {
    await plugins.getByText("Extension connections").waitFor({ state: "visible", timeout: 5_000 });
    await plugins.getByText("Installed extensions haven't registered any connections.").waitFor({ state: "visible", timeout: 5_000 });
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});

void test("Connections renders connections registered by an installed extension", async () => {
  const { context, plugins, pageErrors } = await openConnectionsWithExtension(`
    export function activate(api) {
      api.connections.register({
        id: "apollo",
        title: "Apollo",
        capability: "company enrichment via Apollo API",
        authKind: "api_key",
        secretFields: [{ id: "apiKey", label: "API key", required: true }]
      });
    }
  `, true);
  try {
    const card = plugins.locator(".pi-item-card", { hasText: "Apollo" });
    await card.waitFor({ state: "visible", timeout: 5_000 });
    assert.match((await card.textContent()) ?? "", /company enrichment via Apollo API/u);
    assert.deepEqual(pageErrors.map((error) => error.message), []);
  } finally {
    await context.close();
  }
});
