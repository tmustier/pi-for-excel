import assert from "node:assert/strict";

import {
  chromium,
  type Browser,
  type BrowserContext,
  type BrowserContextOptions,
  type Page,
  type Route,
} from "playwright";
import { createServer, type ViteDevServer } from "vite";

const BRIDGE_URL = "https://localhost:3157";
const LOOPBACK_HOSTNAMES = new Set(["127.0.0.1", "localhost", "[::1]"]);

export interface TaskpaneServer {
  baseUrl: string;
  browser: Browser;
  close(): Promise<void>;
}

interface StartTaskpaneServerOptions {
  token?: string;
}

export type BridgeHandler = (url: URL, route: Route) => Promise<boolean>;

interface OpenBrowserPageOptions {
  path?: string;
  viewport?: BrowserContextOptions["viewport"];
  bridge?: BridgeHandler;
  clientId?: string;
  prepareContext?: (context: BrowserContext) => Promise<void>;
}

type OpenTaskpaneOptions = OpenBrowserPageOptions;

export interface OpenedBrowserPage {
  page: Page;
  context: BrowserContext;
  finish(): Promise<void>;
}

export function isExpectedOfficeUnavailableError(error: Error): boolean {
  return /Office(?:\.js)? (?:is |was )?(?:not ready|not available|unavailable)/i.test(error.message);
}

export async function startTaskpaneServer(
  options: StartTaskpaneServerOptions = {},
): Promise<TaskpaneServer> {
  process.env.VITE_PI_BACKGROUND_VERIFY_URL = BRIDGE_URL;
  process.env.VITE_PI_BACKGROUND_VERIFY_TOKEN = options.token ?? "browser-bridge-token";

  const server: ViteDevServer = await createServer({
    configFile: false,
    root: process.cwd(),
    server: { host: "127.0.0.1", port: 0, strictPort: false },
  });
  await server.listen();

  const localUrl = server.resolvedUrls?.local[0];
  if (!localUrl) {
    await server.close();
    throw new Error("Vite did not expose a local URL");
  }

  let browser: Browser;
  try {
    browser = await chromium.launch({ headless: true });
  } catch (error) {
    await server.close();
    throw error;
  }

  return {
    baseUrl: localUrl,
    browser,
    async close(): Promise<void> {
      await browser.close();
      await server.close();
    },
  };
}

export async function openBrowserPage(
  env: TaskpaneServer,
  options: OpenBrowserPageOptions = {},
): Promise<OpenedBrowserPage> {
  const context = await env.browser.newContext(options.viewport ? { viewport: options.viewport } : {});
  const page = await context.newPage();
  const pageErrors: Error[] = [];

  page.on("pageerror", (error) => {
    if (!isExpectedOfficeUnavailableError(error)) pageErrors.push(error);
  });

  await context.route("**/*", async (route) => {
    const url = new URL(route.request().url());
    if (url.hostname === "localhost" && url.port === "3157") {
      if (await options.bridge?.(url, route)) return;
      if (url.pathname === "/client/register") {
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify({ clientId: options.clientId ?? "idle-browser-client" }),
        });
      } else if (url.pathname === "/client/poll") {
        await new Promise((resolve) => setTimeout(resolve, 1_000));
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ type: "noop" }) });
      } else {
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true }) });
      }
      return;
    }

    if (LOOPBACK_HOSTNAMES.has(url.hostname)) {
      await route.continue();
      return;
    }
    await route.abort("blockedbyclient");
  });

  await options.prepareContext?.(context);

  try {
    await page.goto(`${env.baseUrl}${options.path ?? "src/taskpane.html"}`);
  } catch (error) {
    await context.close();
    throw error;
  }

  return {
    page,
    context,
    async finish(): Promise<void> {
      const messages = pageErrors.map((error) => error.message);
      await context.close();
      assert.deepEqual(messages, [], "unexpected uncaught page errors");
    },
  };
}

export async function openTaskpane(
  env: TaskpaneServer,
  options: OpenTaskpaneOptions = {},
): Promise<OpenedBrowserPage> {
  const opened = await openBrowserPage(env, options);
  try {
    await opened.page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    const welcomeOverlay = opened.page.locator("#pi-welcome-login-overlay");
    await welcomeOverlay.waitFor({ state: "visible", timeout: 10_000 });
    // Dispatch on the backdrop element itself rather than clicking a viewport
    // point: an extension or dialog overlay stacked above the welcome overlay
    // would otherwise receive the click and the welcome overlay would stay.
    // Real backdrop-click dismissal is covered by taskpane-contracts.
    await welcomeOverlay.dispatchEvent("click");
    await welcomeOverlay.waitFor({ state: "detached" });
    return opened;
  } catch (error) {
    await opened.context.close();
    throw error;
  }
}

export async function enterInput(page: Page, text: string): Promise<void> {
  const input = page.locator("pi-input textarea");
  await input.fill(text);
  await input.press("Enter");
}

export async function enterCommand(page: Page, command: string): Promise<void> {
  await enterInput(page, command);
}
