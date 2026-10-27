import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { BrowserContext, Page } from "playwright";

import {
  enterCommand,
  openBrowserPage,
  openTaskpane,
  startTaskpaneServer,
  type TaskpaneServer,
} from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

async function withBrowserPage(
  path: string,
  run: (page: Page) => Promise<void>,
  prepareContext?: (context: BrowserContext) => Promise<void>,
): Promise<void> {
  const opened = await openBrowserPage(env, {
    path,
    ...(prepareContext ? { prepareContext } : {}),
  });
  try {
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env);
  try {
    await run(opened.page);
  } finally {
    await opened.finish();
  }
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

void test("/plugins opens the Plugins view", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/plugins");
    await page.getByRole("heading", { name: "Plugins" }).waitFor({ state: "visible", timeout: 5_000 });
  });
});

async function verifyFailedCommandToast(failureKind: "rejects" | "throws"): Promise<void> {
  await withTaskpane(async (page) => {
    const sentinel = `browser-command-${failureKind}-sentinel`;
    await page.evaluate(async ({ kind, message }: { kind: "rejects" | "throws"; message: string }) => {
      const { commandRegistry } = await import("/src/commands/types.ts") as typeof import("../../src/commands/types.ts");
      commandRegistry.register({
        name: `browser-${kind}`,
        description: "Browser failure contract",
        source: "extension",
        execute: kind === "rejects"
          ? () => Promise.reject(new Error(message))
          : () => { throw new Error(message); },
      });
    }, { kind: failureKind, message: sentinel });

    await enterCommand(page, `/browser-${failureKind}`);
    const toast = page.locator("#pi-toast .pi-toast__message").last();
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
    // Stepping out of the field keeps focus inside the dialog, not on <body>.
    assert.equal(await settings.evaluate((overlay) => overlay.contains(document.activeElement)), true);
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

void test("typing /backup in the composer and pressing Enter runs the manual backup command", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/backup");

    const toast = page.locator("#pi-toast .pi-toast__message").last();
    await toast.waitFor({ state: "visible", timeout: 5_000 });
    assert.match(await toast.innerText(), /Backup command failed: .*unavailable/i);
  });
});

void test("submitInput uses the composer path for commands and model messages", async () => {
  const commands = [
    { id: "slash-input", type: "submitInput", payload: { text: "/backup", waitForIdle: true } },
    { id: "plain-input", type: "submitInput", payload: { text: "browser bridge model message", waitForIdle: false } },
  ];
  const results = new Map<string, { ok: boolean }>();
  const pendingResultIds = ["slash-input", "plain-input"];
  let allowCommands = false;
  let allowPlainInput = false;
  let resolveResults = (): void => {};
  const allResults = new Promise<void>((resolve) => {
    resolveResults = resolve;
  });

  await withBrowserPage("src/taskpane.html", async (page) => {
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    const welcomeOverlay = page.locator("#pi-welcome-login-overlay");
    await welcomeOverlay.waitFor({ state: "visible", timeout: 10_000 });
    await welcomeOverlay.click({ position: { x: 2, y: 2 } });
    await welcomeOverlay.waitFor({ state: "detached" });
    allowCommands = true;

    const toast = page.locator("#pi-toast .pi-toast__message").last();
    await toast.waitFor({ state: "visible", timeout: 5_000 });
    assert.match(await toast.innerText(), /Backup command failed: .*unavailable/i);
    allowPlainInput = true;
    await page.getByText("browser bridge model message", { exact: true }).waitFor({ state: "visible", timeout: 5_000 });
    await waitForBrowserSignal(allResults, "Timed out waiting for submitInput bridge results");
    assert.deepEqual([...results.values()], [{ ok: true }, { ok: true }]);
  }, async (context) => {
    await context.route("https://localhost:3157/**", async (route) => {
      const url = new URL(route.request().url());
      if (url.pathname === "/client/register") {
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ clientId: "browser-client" }) });
        return;
      }
      if (url.pathname === "/client/poll") {
        const nextCommand = commands[0];
        const canSendNext = allowCommands && (nextCommand?.id !== "plain-input" || allowPlainInput);
        const command = canSendNext ? commands.shift() ?? { type: "noop" } : { type: "noop" };
        await route.fulfill({ contentType: "application/json", body: JSON.stringify(command) });
        return;
      }
      if (url.pathname === "/client/result") {
        const commandId = pendingResultIds.shift();
        if (commandId) {
          results.set(commandId, { ok: /"ok":true/u.test(route.request().postData() ?? "") });
          if (results.size === 2) resolveResults();
        }
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true }) });
        return;
      }
      await route.abort("blockedbyclient");
    });
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

    const toast = page.locator("#pi-toast .pi-toast__message").last();
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

void test("first-run disclosure is a dismissible banner that can open Settings", async () => {
  await withBrowserPage("src/ui-gallery.html", async (page) => {
    await page.evaluate(`
      (async () => {
        localStorage.removeItem("pi.onboarding.disclosure.acknowledged");
        const { createDisclosureBar } = await import("/src/ui/disclosure-bar.ts");
        window.__settingsOpened = 0;
        const bar = createDisclosureBar({ providerCount: 1, onOpenSettings: () => { window.__settingsOpened += 1; } });
        if (bar) document.body.appendChild(bar.root);
      })()
    `);

    const banner = page.locator('body > .pi-banner[data-banner-id="disclosure"]');
    await banner.waitFor({ state: "visible" });
    assert.equal(await banner.getAttribute("role"), "status");
    await banner.getByRole("button", { name: "Review in Settings" }).click();
    await banner.waitFor({ state: "detached" });
    assert.equal(await page.evaluate("window.__settingsOpened"), 1);

    // Acknowledged for good: a second taskpane boot does not show it again.
    assert.equal(await page.evaluate(`
      (async () => {
        const { createDisclosureBar } = await import("/src/ui/disclosure-bar.ts");
        return createDisclosureBar({ providerCount: 1, onOpenSettings: () => {} }) === null;
      })()
    `), true);
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
    // An inactive tab's × appears on hover.
    await page.locator(".pi-session-tab").first().hover();
    await closeButtons.first().click();
    await page.waitForFunction(() => document.querySelectorAll('.pi-session-tab__close').length === 0);

    await enterCommand(page, "/resume");
    const resume = page.locator("#pi-resume-overlay");
    await resume.getByRole("heading", { name: "Recently closed" }).waitFor({ state: "visible", timeout: 5_000 });
    await resume.locator('[data-resume-section="recently-closed"] .pi-resume-item').waitFor({ state: "visible" });

    // Tab from the card itself must be handled by the trap: the browser's default
    // from a tabindex=-1 card leaves the add-in frame inside Office.
    const tabFromCard = await page.evaluate(() => {
      const card = document.querySelector<HTMLElement>("#pi-resume-overlay .pi-overlay-card, #pi-resume-overlay [tabindex=\"-1\"]");
      card?.focus();
      const event = new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true });
      (document.activeElement ?? document.body).dispatchEvent(event);
      const active = document.activeElement;
      return { prevented: event.defaultPrevented, inside: card?.contains(active) === true && active !== card };
    });
    assert.deepEqual(tabFromCard, { prevented: true, inside: true });
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

void test("Settings back/close rule: Escape closes on the entry page and steps back from a nested one", async () => {
  await withTaskpane(async (page) => {
    const settings = page.locator("#pi-settings-overlay");

    // Opened straight at Backups: Escape closes instead of landing on the Settings root.
    await enterCommand(page, "/history");
    await settings.getByRole("heading", { name: "Backups" }).waitFor({ state: "visible", timeout: 5_000 });
    await page.keyboard.press("Escape");
    await settings.waitFor({ state: "detached" });

    // Reached from the Settings root: Escape goes back to the root, then closes.
    await enterCommand(page, "/settings");
    await settings.getByRole("button", { name: /^Backups/ }).click();
    await settings.getByRole("heading", { name: "Backups" }).waitFor({ state: "visible" });
    await page.keyboard.press("Escape");
    await settings.getByRole("heading", { name: "Settings" }).waitFor({ state: "visible" });
    await page.keyboard.press("Escape");
    await settings.waitFor({ state: "detached" });
  });
});

void test("proxy banner warns only when the enabled proxy is down, and can be dismissed", async () => {
  await withTaskpane(async (page) => {
    const banner = page.locator('#pi-notices .pi-banner[data-banner-id="proxy"]');
    const setProxyState = (state: string) => page.evaluate((next) => {
      document.dispatchEvent(new CustomEvent("pi:proxy-state-changed", { detail: { state: next } }));
    }, state);
    const setProxyEnabled = (enabled: boolean) => page.evaluate(async (next) => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      await getAppStorage().settings.set("proxy.enabled", next);
    }, enabled);

    await setProxyEnabled(false);
    await setProxyState("not-detected");
    await page.waitForTimeout(200);
    assert.equal(await banner.count(), 0);

    await setProxyEnabled(true);
    await setProxyState("not-detected");
    await banner.waitFor({ state: "visible" });
    assert.match(await banner.innerText(), /proxy is turned on but isn’t running/);

    await banner.getByRole("button", { name: "Dismiss" }).click();
    await banner.waitFor({ state: "detached" });
    await setProxyState("not-detected");
    await page.waitForTimeout(200);
    assert.equal(await banner.count(), 0, "a dismissed warning stays hidden until the proxy recovers");

    await setProxyState("detected");
    await setProxyState("not-detected");
    await banner.waitFor({ state: "visible" });
    await setProxyState("detected");
    await banner.waitFor({ state: "detached" });
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
