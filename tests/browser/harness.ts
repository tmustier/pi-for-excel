import assert from "node:assert/strict";

import {
  chromium,
  type Browser,
  type BrowserContext,
  type BrowserContextOptions,
  type Page,
  type Route,
} from "playwright";
import { build, preview, type PreviewServer } from "vite";

import { mkdtemp, readdir, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const BRIDGE_URL = "https://localhost:3157";
const BROWSER_MODULE_URL = /["'](\/(?:src\/[^"']+\.ts|node_modules\/[^"']+\.js))["']/gu;
const BUNDLED_MODULE_PATH = /^\/(?:src\/.+\.ts|node_modules\/.+\.js)$/u;
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

export async function connectFixtureProvider(page: Page, provider = "mistral", key = "fixture-key"): Promise<void> {
  await page.evaluate(async ({ provider, key }) => {
    const path = "/src/storage/local/app-storage.ts";
    const { getAppStorage } = await import(/* @vite-ignore */ path) as typeof import("../../src/storage/local/app-storage.ts");
    await getAppStorage().providerKeys.set(provider, key);
    document.dispatchEvent(new Event("pi:providers-changed"));
  }, { provider, key });
  await page.waitForFunction(async (provider) => {
    const path = "/src/ui/pi-sidebar.ts";
    const { PiSidebar } = await import(/* @vite-ignore */ path) as typeof import("../../src/ui/pi-sidebar.ts");
    const sidebar = document.querySelector("pi-sidebar");
    return sidebar instanceof PiSidebar && sidebar.agent?.state.model.provider === provider;
  }, provider);
}

export function isExpectedOfficeUnavailableError(error: Error): boolean {
  return /Office(?:\.js)? (?:is |was )?(?:not ready|not available|unavailable)/i.test(error.message);
}

export async function startTaskpaneServer(
  options: StartTaskpaneServerOptions = {},
): Promise<TaskpaneServer> {
  process.env.VITE_PI_BACKGROUND_VERIFY_URL = BRIDGE_URL;
  process.env.VITE_PI_BACKGROUND_VERIFY_TOKEN = options.token ?? "browser-bridge-token";

  const outDir = await mkdtemp(path.join(os.tmpdir(), "pi-browser-test-"));
  let browser: Browser;
  let server: PreviewServer;
  try {
    [browser, server] = await Promise.all([
      chromium.launch({ headless: true }),
      buildTaskpane(outDir).then(() => serveTaskpane(outDir)),
    ]);
  } catch (error) {
    await rm(outDir, { recursive: true, force: true });
    throw error;
  }

  const localUrl = server.resolvedUrls?.local[0];
  if (!localUrl) {
    await browser.close();
    await server.close();
    await rm(outDir, { recursive: true, force: true });
    throw new Error("Vite did not expose a local URL");
  }

  return {
    baseUrl: localUrl,
    browser,
    async close(): Promise<void> {
      await browser.close();
      await server.close();
      await rm(outDir, { recursive: true, force: true });
    },
  };
}

async function sourceModulesImportedByTests(): Promise<Record<string, string>> {
  const dir = import.meta.dirname;
  const entries: Record<string, string> = {};
  for (const name of await readdir(dir)) {
    if (!name.endsWith(".ts")) continue;
    const source = await readFile(path.join(dir, name), "utf8");
    for (const match of source.matchAll(BROWSER_MODULE_URL)) {
      const file = match[1].slice(1);
      entries[file.slice(0, -path.extname(file).length)] = file;
    }
  }
  return entries;
}

async function buildTaskpane(outDir: string): Promise<void> {
  process.env.NODE_ENV ??= "development";
  await build({
    root: process.cwd(),
    mode: "development",
    logLevel: "error",
    build: {
      outDir,
      emptyOutDir: true,
      minify: false,
      rollupOptions: {
        input: {
          taskpane: "src/taskpane.html",
          "ui-gallery": "src/ui-gallery.html",
          ...(await sourceModulesImportedByTests()),
        },
        preserveEntrySignatures: "allow-extension",
        output: {
          entryFileNames: (chunk) => (chunk.name.startsWith("src/") ? "[name].ts" : "[name].js"),
        },
      },
    },
  });
}

async function serveTaskpane(outDir: string): Promise<PreviewServer> {
  const root = process.cwd();
  return preview({
    configFile: false,
    root,
    logLevel: "error",
    build: { outDir },
    preview: { host: "127.0.0.1", port: 0, strictPort: false, open: false },
    plugins: [{
      name: "pi-browser-test-sources",
      configurePreviewServer(server) {
        server.middlewares.use((req, res, next) => {
          const url = new URL(req.url ?? "/", "http://localhost");
          const bundled = BUNDLED_MODULE_PATH.test(url.pathname);
          if (!bundled && !(url.pathname.startsWith("/src/") && url.pathname.endsWith(".css"))) return next();
          void readFile(path.join(bundled ? outDir : root, url.pathname)).then((body) => {
            res.setHeader("Content-Type", bundled ? "text/javascript; charset=utf-8" : "text/css; charset=utf-8");
            res.end(body);
          }, next);
        });
      },
    }],
  });
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
    await welcomeOverlay.dispatchEvent("click");
    await welcomeOverlay.waitFor({ state: "detached" });
    return opened;
  } catch (error) {
    await opened.context.close();
    throw error;
  }
}

export async function waitForBrowserSignal(signal: Promise<void>, timeoutMessage: string): Promise<void> {
  let timeoutId: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<void>((_resolve, reject) => {
    timeoutId = setTimeout(() => reject(new Error(timeoutMessage)), 10_000);
  });

  try {
    await Promise.race([signal, timeout]);
  } finally {
    if (timeoutId !== undefined) {
      clearTimeout(timeoutId);
    }
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
