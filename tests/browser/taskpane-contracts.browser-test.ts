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

async function withBrowserPage(
  path: string,
  run: (page: Page) => Promise<void>,
  prepareContext?: (context: BrowserContext) => Promise<void>,
): Promise<void> {
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
  await prepareContext?.(context);

  try {
    await page.goto(`${baseUrl}${path}`);
    await run(page);
    assert.deepEqual(pageErrors.map((error) => error.message), [], "unexpected uncaught page errors");
  } finally {
    await context.close();
  }
}

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  await withBrowserPage("src/taskpane.html", async (page) => {
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });

    const welcomeOverlay = page.locator("#pi-welcome-login-overlay");
    await welcomeOverlay.waitFor({ state: "visible", timeout: 10_000 });
    await welcomeOverlay.click({ position: { x: 2, y: 2 } });
    await welcomeOverlay.waitFor({ state: "detached" });

    await run(page);
  });
}

async function enterCommand(page: Page, command: string): Promise<void> {
  const input = page.locator("pi-input textarea");
  await input.fill(command);
  await input.press("Enter");
}

async function openUtilitiesMenu(page: Page): Promise<void> {
  await page.getByRole("button", { name: "Settings and tools" }).click();
}

async function waitForBrowserSignal(signal: Promise<void>, timeoutMessage: string): Promise<void> {
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

void test("/files opens the Files view", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/files");
    await page.getByRole("heading", { name: "Files" }).waitFor({ state: "visible" });
  });
});

void test("/plugins exposes the MCP server setup flow", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/plugins");
    await page.getByRole("heading", { name: "Connections" }).waitFor({ state: "visible", timeout: 5_000 });
    await page.getByText("MCP servers", { exact: true }).waitFor({ state: "visible" });
    await page.getByRole("button", { name: "+ Add server" }).click();
    await page.getByPlaceholder("https://server-url/rpc").waitFor({ state: "visible" });
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

void test("closing Settings restores focus to the chat input", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.focus();
    await openUtilitiesMenu(page);
    await page.getByRole("menuitem", { name: "Settings" }).click();

    const settings = page.locator("#pi-settings-overlay");
    await page.getByRole("heading", { name: "Settings" }).waitFor({ state: "visible", timeout: 5_000 });
    await settings.getByRole("button", { name: "Close Settings" }).click();
    await settings.waitFor({ state: "detached" });
    await page.waitForFunction(() => document.activeElement?.tagName === "TEXTAREA", undefined, { timeout: 3_000 });
    assert.equal(await input.evaluate((element) => element === document.activeElement), true);
  });
});

void test("Settings ignores card clicks and closes on a backdrop click", async () => {
  await withTaskpane(async (page) => {
    await openUtilitiesMenu(page);
    await page.getByRole("menuitem", { name: "Settings" }).click();

    const settings = page.locator("#pi-settings-overlay");
    await settings.locator(".pi-set-shell").click();
    await settings.waitFor({ state: "visible" });

    await settings.click({ position: { x: 2, y: 2 } });
    await settings.waitFor({ state: "detached" });
  });
});

void test("Escape preserves keyboard entry and closes only the topmost dialog", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/settings");
    const settings = page.locator("#pi-settings-overlay");
    await page.getByRole("heading", { name: "Settings" }).waitFor({ state: "visible", timeout: 5_000 });
    await settings.getByRole("button", { name: "Custom gateway" }).click();
    await page.getByRole("heading", { name: "Custom OpenAI-compatible gateway" }).waitFor({ state: "visible" });

    const nameInput = page.getByPlaceholder("Gateway name (optional)");
    await nameInput.fill("Browser gateway");
    await nameInput.press("Escape");
    assert.equal(await nameInput.inputValue(), "Browser gateway");
    assert.equal(await nameInput.evaluate((element) => element === document.activeElement), false);
    await settings.waitFor({ state: "visible" });

    await page.getByPlaceholder("https://your-gateway.example.com/v1").fill("https://gateway.example.com/v1");
    await page.getByPlaceholder("model-id").fill("browser-model");
    await page.getByRole("button", { name: "Save gateway" }).click();
    await settings.getByText("Browser gateway", { exact: true }).waitFor({ state: "visible" });
    await settings.getByRole("button", { name: "Delete" }).click();

    const confirmation = page.locator("#pi-confirm-dialog-overlay");
    await confirmation.waitFor({ state: "visible" });
    await page.keyboard.press("Escape");
    await confirmation.waitFor({ state: "detached", timeout: 3_000 });
    await settings.waitFor({ state: "visible" });
    await page.getByRole("heading", { name: "Custom OpenAI-compatible gateway" }).waitFor({ state: "visible" });

    await page.keyboard.press("Escape");
    await page.getByRole("heading", { name: "Settings" }).waitFor({ state: "visible" });
    await page.keyboard.press("Escape");
    await settings.waitFor({ state: "detached" });
  });
});

void test("copy fallback selects the full command in Chromium", async () => {
  await withBrowserPage("src/ui-gallery.html", async (page) => {
    const row = page.locator(".pi-command-copy").filter({ hasText: "npx pi-for-excel-proxy" });
    await row.waitFor({ state: "visible", timeout: 10_000 });
    await row.getByRole("button", { name: "Copy command" }).click();

    await page.waitForFunction(() => window.getSelection()?.toString() === "npx pi-for-excel-proxy");
    assert.equal(await page.evaluate(() => window.getSelection()?.toString()), "npx pi-for-excel-proxy");
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

void test("Extensions exposes Connections, Plugins, and Skills navigation", async () => {
  await withTaskpane(async (page) => {
    await openUtilitiesMenu(page);
    await page.getByRole("menuitem", { name: "Extensions" }).click();
    const settings = page.locator("#pi-settings-overlay");
    await settings.getByRole("heading", { name: "Connections" }).waitFor({ state: "visible", timeout: 5_000 });
    await settings.getByRole("button", { name: "Back" }).click();

    for (const pageName of ["Connections", "Plugins", "Skills"]) {
      const navigation = settings.getByRole("button", { name: new RegExp(`^${pageName}`) });
      await navigation.waitFor({ state: "visible" });
      await navigation.click();
      await settings.getByRole("heading", { name: pageName }).waitFor({ state: "visible" });
      await settings.getByRole("button", { name: "Back" }).click();
    }
  });
});

void test("Settings Backups exposes manual backup and reports browser-host unavailability", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(`
      (async () => {
        const { getSettingsPagesDependencies } = await import("/src/commands/builtins/settings-pages/dependencies.ts");
        const backups = getSettingsPagesDependencies().backups;
        if (!backups) throw new Error("Backups dependencies were not configured");
        backups.loadCheckpoints = () => Promise.resolve([{
          id: "browser-checkpoint",
          at: Date.now(),
          toolName: "write_cells",
          address: "Sheet1!A1",
          changedCount: 1,
        }]);
      })()
    `);

    await openUtilitiesMenu(page);
    await page.getByRole("menuitem", { name: "Settings" }).click();
    const settings = page.locator("#pi-settings-overlay");
    await settings.getByRole("button", { name: /^Backups/ }).click();
    await settings.getByRole("heading", { name: "Backups" }).waitFor({ state: "visible" });

    const manualBackup = settings.getByRole("button", { name: "Download backup" });
    await manualBackup.waitFor({ state: "visible" });
    await manualBackup.click();

    const toast = page.locator("#pi-toast.visible .pi-toast__message");
    await toast.waitFor({ state: "visible", timeout: 5_000 });
    assert.match(await toast.innerText(), /Backup failed: .*unavailable/i);
  });
});

void test("local-service probes populate the first runtime capabilities", async () => {
  let releaseProbes = (): void => {};
  const probesReleased = new Promise<void>((resolve) => {
    releaseProbes = resolve;
  });
  let markProbeStarted = (): void => {};
  const probeStarted = new Promise<void>((resolve) => {
    markProbeStarted = resolve;
  });

  await withBrowserPage("src/taskpane.html", async (page) => {
    try {
      await waitForBrowserSignal(probeStarted, "Timed out waiting for a local-service health probe to start");
      assert.equal(await page.locator("pi-input textarea").count(), 0);
    } finally {
      releaseProbes();
    }

    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    const systemPrompt = await page.evaluate(`document.querySelector("pi-sidebar")?.agent?.state?.systemPrompt ?? ""`);
    assert.match(systemPrompt, /python 3\.browser\.12/);
    assert.match(systemPrompt, /tmux browser-sentinel/);
  }, async (context) => {
    await context.route("https://localhost:3340/health", async (route) => {
      markProbeStarted();
      await waitForBrowserSignal(probesReleased, "Timed out waiting to release the Python health probe");
      await route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({
          ok: true,
          python: { available: true, version: "3.browser.12" },
          libreoffice: { available: true },
        }),
      });
    });
    await context.route("https://localhost:3341/health", async (route) => {
      markProbeStarted();
      await waitForBrowserSignal(probesReleased, "Timed out waiting to release the tmux health probe");
      await route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({ ok: true, tmuxVersion: "tmux browser-sentinel", sessions: 2 }),
      });
    });
  });
});

void test("disclosure customization renders operable shared toggles", async () => {
  await withBrowserPage("src/ui-gallery.html", async (page) => {
    await page.evaluate(`
      (async () => {
        const { createDisclosureBar } = await import("/src/ui/disclosure-bar.ts");
        const firstBar = createDisclosureBar({ providerCount: 1 });
        const secondBar = createDisclosureBar({ providerCount: 1 });
        if (firstBar) document.body.appendChild(firstBar);
        if (secondBar) document.body.appendChild(secondBar);
      })()
    `);

    const bars = page.locator(".pi-disclosure-bar");
    assert.equal(await bars.count(), 2);
    const firstBar = bars.nth(0);
    const secondBar = bars.nth(1);
    const customize = firstBar.locator(".pi-disclosure-bar__link");
    const secondCustomize = secondBar.locator(".pi-disclosure-bar__link");
    const pickerId = await customize.getAttribute("aria-controls");
    const secondPickerId = await secondCustomize.getAttribute("aria-controls");
    assert.ok(pickerId);
    assert.ok(secondPickerId);
    assert.notEqual(pickerId, secondPickerId);
    assert.equal(await firstBar.locator(".pi-disclosure-picker").getAttribute("id"), pickerId);
    assert.equal(await secondBar.locator(".pi-disclosure-picker").getAttribute("id"), secondPickerId);

    assert.equal(await customize.getAttribute("aria-expanded"), "false");
    await firstBar.locator(`#${pickerId}`).waitFor({ state: "hidden" });
    await customize.click();
    assert.equal(await customize.getAttribute("aria-expanded"), "true");
    await firstBar.locator(`#${pickerId}`).waitFor({ state: "visible" });

    const webSearchRow = firstBar.locator(".pi-toggle-row").filter({ hasText: "Web search" });
    const checkbox = webSearchRow.locator('input[type="checkbox"]');
    await checkbox.waitFor({ state: "attached" });
    assert.equal(await checkbox.isChecked(), true);
    await webSearchRow.locator("label.pi-toggle").click();
    assert.equal(await checkbox.isChecked(), false);
  });
});

void test("paperclip, context disclosure, and non-streaming Escape remain operable", async () => {
  await withTaskpane(async (page) => {
    await page.getByRole("button", { name: "Browse files" }).click();
    await page.getByRole("heading", { name: "Files" }).waitFor({ state: "visible", timeout: 5_000 });
    await page.getByRole("button", { name: "Close Files" }).click();

    await page.evaluate(`
      (async () => {
        localStorage.setItem("pi-excel.debug", "1");
        const { getPayloadStats } = await import("/src/auth/stream-proxy.ts");
        Object.assign(getPayloadStats(), { calls: 1, messageCount: 1, messageChars: 7 });
        const sidebar = document.querySelector("pi-sidebar");
        sidebar.agent.state.messages.push({ role: "user", content: "browser" });
        sidebar.requestUpdate();
        document.dispatchEvent(new Event("pi:debug-changed"));
      })()
    `);

    const contextHeader = page.locator(".pi-context-pill__header");
    await contextHeader.waitFor({ state: "visible" });
    assert.equal(await contextHeader.getAttribute("aria-expanded"), "false");
    const controlledBodyId = await contextHeader.getAttribute("aria-controls");
    assert.ok(controlledBodyId);
    assert.equal(await page.locator(`#${controlledBodyId}`).count(), 0);
    await contextHeader.click();
    assert.equal(await contextHeader.getAttribute("aria-expanded"), "true");
    await page.locator(`#${controlledBodyId}`).waitFor({ state: "visible" });
    await contextHeader.click();
    assert.equal(await contextHeader.getAttribute("aria-expanded"), "false");
    assert.equal(await page.locator(`#${controlledBodyId}`).count(), 0);

    await page.evaluate(`
      const slot = document.querySelector("#pi-widget-slot");
      const widget = document.createElement("button");
      widget.textContent = "Browser widget";
      slot.appendChild(widget);
      slot.style.display = "block";
    `);
    const input = page.locator("pi-input textarea");
    await input.fill("preserve this draft");
    await input.press("Escape");
    assert.equal(await input.inputValue(), "preserve this draft");
    assert.equal(await input.evaluate((element) => element === document.activeElement), false);
  });
});

void test("recently closed tabs and experimental settings are reachable by commands", async () => {
  await withTaskpane(async (page) => {
    await page.getByRole("button", { name: "New tab" }).click();
    const closeButtons = page.getByRole("button", { name: "Close tab" });
    await closeButtons.first().waitFor({ state: "visible" });
    assert.equal(await closeButtons.count(), 2);
    await closeButtons.first().click();
    await page.waitForFunction(() => document.querySelectorAll('.pi-session-tab__close').length === 0);

    await enterCommand(page, "/resume");
    const resume = page.locator("#pi-resume-overlay");
    await resume.getByRole("heading", { name: "Recently closed" }).waitFor({ state: "visible", timeout: 5_000 });
    await resume.locator('[data-resume-section="recently-closed"] .pi-resume-item').waitFor({ state: "visible" });
    await resume.getByRole("button", { name: "Close resume session" }).click();

    await enterCommand(page, "/experimental");
    const settings = page.locator("#pi-settings-overlay");
    await settings.getByRole("heading", { name: "Experimental features", level: 2 }).waitFor({ state: "visible", timeout: 5_000 });
    await settings.getByText("Dark mode", { exact: true }).waitFor({ state: "visible" });
  });
});

void test("Settings guards unsaved navigation and exposes shared behavior controls", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/settings");
    const settings = page.locator("#pi-settings-overlay");
    await settings.getByText("Auto mode", { exact: true }).waitFor({ state: "visible", timeout: 5_000 });
    await settings.getByText("Fork model switch into new tab", { exact: true }).waitFor({ state: "visible" });
    await settings.getByRole("button", { name: /Rules & conventions/ }).click();

    const rules = page.getByPlaceholder(/Your preferences and habits/);
    await rules.fill("Always show units");
    await settings.getByRole("button", { name: "Back" }).click();

    const confirmation = page.locator("#pi-confirm-dialog-overlay");
    await confirmation.getByRole("heading", { name: "Discard changes?" }).waitFor({ state: "visible" });
    await confirmation.getByRole("button", { name: "Keep editing" }).filter({ hasText: "Keep editing" }).click();
    await confirmation.waitFor({ state: "detached" });
    assert.equal(await rules.inputValue(), "Always show units");
    await settings.getByRole("button", { name: "Back" }).click();
    await page.getByRole("button", { name: "Discard" }).click();
    await settings.getByRole("heading", { name: "Settings" }).waitFor({ state: "visible" });
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
