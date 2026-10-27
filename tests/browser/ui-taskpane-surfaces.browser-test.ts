import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { BrowserContext, Page } from "playwright";

import { enterCommand, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

async function withTaskpane(run: (page: Page, context: BrowserContext) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, { viewport: { width: 420, height: 720 }, clientId: "ui-taskpane-browser" });
  try {
    await opened.page.evaluate(() => {
      window.addEventListener("beforeunload", () => {
        document.body.dataset.browserPageErrorsChecked = "true";
      });
    });
    await run(opened.page, opened.context);
  } finally {
    await opened.finish();
  }
}

void test("open tab titles survive a taskpane reload", async () => {
  await withTaskpane(async (page) => {
    await page.getByRole("button", { name: "New tab" }).click();
    await page.locator(".pi-session-tab__title").filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });
    await page.waitForTimeout(300);

    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });

    assert.deepEqual(await page.locator(".pi-session-tab__title").allTextContents(), ["Chat 1", "Chat 2"]);
  });
});

void test("renaming a tab uses the in-pane dialog, focused on the title field", async () => {
  await withTaskpane(async (page) => {
    page.once("dialog", () => assert.fail("tab rename must not use a native browser dialog"));
    await page.locator(".pi-session-tab__main").dblclick();
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    const input = dialog.getByRole("textbox");
    await input.waitFor({ state: "visible" });
    assert.equal(await input.evaluate((node) => node === document.activeElement), true);
    await page.keyboard.press("ControlOrMeta+A");
    await page.keyboard.type("Revenue Model");
    await page.keyboard.press("Enter");
    await dialog.waitFor({ state: "detached" });

    await page.locator(".pi-session-tab__title").filter({ hasText: "Revenue Model" }).waitFor({ state: "visible" });
    assert.equal(await page.locator(".pi-session-tab__title").innerText(), "Revenue Model");
  });
});

void test("a busy tab's model picker lists models but cannot pick one, and unlocks when the turn ends", async () => {
  await withTaskpane(async (page) => {
    // A connected provider, so the picker has models to list.
    await page.evaluate(async () => {
      // Resolves from the page (/src/taskpane.html) to the module the app itself loaded.
      const storage = await import("../../src/storage/local/app-storage.ts");
      await storage.getAppStorage().providerKeys.set("openai", "sk-test");
    });
    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    const modelBefore = await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar)) throw new Error("sidebar missing");
      // The sidebar's agent is the active tab's Agent: a stream that never answers keeps the turn running.
      const agent = sidebar.agent as {
        state: { model: { id: string } };
        getApiKey: () => Promise<string>;
        streamFunction: (model: unknown, context: unknown, options?: { signal?: AbortSignal }) => Promise<never>;
      };
      agent.getApiKey = () => Promise.resolve("test-key");
      // Never answers; Stop (abort) ends the turn.
      agent.streamFunction = (_model, _context, options) => new Promise<never>((_resolve, reject) => {
        options?.signal?.addEventListener("abort", () => reject(new Error("aborted")));
      });
      return agent.state.model.id;
    });
    const welcome = page.locator("#pi-welcome-login-overlay");
    if (await welcome.count()) {
      await welcome.dispatchEvent("click");
      await welcome.waitFor({ state: "detached" });
    }
    await enterCommand(page, "Keep working");
    await page.locator(".pi-input-btn--abort").waitFor({ state: "visible" });

    await enterCommand(page, "/model");
    const picker = page.locator("#pi-model-selector-overlay");
    await picker.getByRole("status").filter({ hasText: "Pi is working in this tab" }).waitFor({ state: "visible" });
    const rows = picker.locator(".pi-model-selector-item");
    await rows.nth(1).waitFor({ state: "visible" });
    assert.equal(await picker.locator(".pi-model-selector-item:not([aria-disabled=true])").count(), 0);

    // force: Playwright will not click an aria-disabled row; the test proves clicking it does nothing.
    await rows.nth(1).click({ force: true });
    await page.keyboard.press("Enter");
    assert.equal(await picker.isVisible(), true);
    assert.equal(await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      return sidebar && "agent" in sidebar ? (sidebar.agent as { state: { model: { id: string } } }).state.model.id : null;
    }), modelBefore);

    // The turn ends while the picker stays open: the rows unlock without reopening it.
    await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (sidebar && "agent" in sidebar) (sidebar.agent as { abort: () => void }).abort();
    });
    await picker.getByRole("status").waitFor({ state: "hidden" });
    await page.waitForFunction(() => document.querySelectorAll(
      "#pi-model-selector-overlay .pi-model-selector-item[aria-disabled=true]",
    ).length === 0);
    const pickedId = await rows.nth(1).locator(".pi-model-selector-item-id").innerText();
    await rows.nth(1).click();
    await picker.waitFor({ state: "detached" });
    await page.waitForFunction((id) => {
      const sidebar = document.querySelector("pi-sidebar");
      return sidebar !== null && "agent" in sidebar
        && (sidebar.agent as { state: { model: { id: string } } }).state.model.id === id;
    }, pickedId);
  });
});

void test("a tab saved on a provider that is no longer connected reloads on a connected model", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      // Resolves from the page (/src/taskpane.html) to the module the app itself loaded.
      const storage = await import("../../src/storage/local/app-storage.ts");
      await storage.getAppStorage().providerKeys.set("openai", "sk-test");
    });
    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    const activeModel = () => page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar)) return null;
      return (sidebar.agent as { state: { model: { provider: string } } }).state.model.provider;
    });
    await page.waitForFunction(() => {
      const sidebar = document.querySelector("pi-sidebar");
      return sidebar !== null && "agent" in sidebar
        && (sidebar.agent as { state: { model: { provider: string } } }).state.model.provider === "openai";
    });

    // The tab's saved model belongs to Google, which has no credentials here.
    await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar)) throw new Error("sidebar missing");
      const agent = sidebar.agent as { state: { model: Record<string, unknown> } };
      agent.state.model = { ...agent.state.model, provider: "google", id: "gemini-2.5-flash-lite" };
    });
    // Renaming saves the session, model included.
    await page.locator(".pi-session-tab.is-active .pi-session-tab__main").dblclick();
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    await dialog.locator("input").fill("Gemini tab");
    await dialog.getByRole("button", { name: "Rename" }).click();
    await dialog.waitFor({ state: "detached" });
    await page.waitForTimeout(300);

    await page.reload();
    await page.locator(".pi-session-tab__title").filter({ hasText: "Gemini tab" }).waitFor({ state: "visible", timeout: 20_000 });
    await page.waitForFunction(() => {
      const sidebar = document.querySelector("pi-sidebar");
      return sidebar !== null && "agent" in sidebar
        && (sidebar.agent as { state: { model: { provider: string } } }).state.model.provider !== "google";
    }, undefined, { timeout: 10_000 }).catch(() => undefined);
    assert.equal(await activeModel(), "openai");
  });
});

void test("the Files detail view exposes writable actions", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(`
      (async () => {
        const { getFilesWorkspace } = await import("/src/files/workspace.ts");
        await getFilesWorkspace().importFiles([
          new File(["quarterly notes"], "plan.md", { type: "text/markdown", lastModified: 1 }),
        ]);
      })()
    `);
    await enterCommand(page, "/files");
    await page.getByText("plan.md", { exact: true }).click();

    const actionButtons = page.locator(".pi-files-detail-actions button");
    await actionButtons.last().waitFor({ state: "visible" });
    assert.deepEqual(
      await actionButtons.allTextContents(),
      ["Open ↗", "Download", "Rename", "Delete"],
    );
  });
});

void test("Files rename keeps the original extension when it is omitted", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(`
      (async () => {
        const { getFilesWorkspace } = await import("/src/files/workspace.ts");
        await getFilesWorkspace().importFiles([
          new File(["quarterly notes"], "plan.md", { type: "text/markdown", lastModified: 1 }),
        ]);
      })()
    `);
    await enterCommand(page, "/files");
    await page.getByText("plan.md", { exact: true }).click();
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => {
      button.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    const renameDialog = page.locator("#pi-text-input-dialog-overlay");
    await renameDialog.getByRole("textbox").fill("archive/final-plan");
    await renameDialog.getByRole("button", { name: "Rename" }).click();

    await page.getByRole("heading", { name: "final-plan.md" }).waitFor({ state: "visible" });
    assert.equal(await page.getByRole("heading", { name: "final-plan.md" }).innerText(), "final-plan.md");
  });
});

void test("tool Markdown renders without its YAML frontmatter", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(`
      (async () => {
        const sidebar = document.querySelector("pi-sidebar");
        sidebar.agent.state.messages.push(
          {
            role: "assistant",
            content: [{ type: "toolCall", id: "markdown-tool", name: "skills", arguments: { action: "read", name: "browser-fixture" } }],
            api: "openai-responses",
            provider: "openai",
            model: "browser-fixture",
            usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
            stopReason: "toolUse",
            timestamp: 1,
          },
          {
            role: "toolResult",
            toolCallId: "markdown-tool",
            toolName: "skills",
            content: [{ type: "text", text: "---\\nname: browser-sentinel\\ndescription: Browser fixture\\n---\\n# Visible guidance\\nBody content." }],
            isError: false,
            timestamp: 2,
          },
        );
        sidebar.syncFromAgent();
        await sidebar.updateComplete;
        await document.querySelector("message-list").updateComplete;
        await document.querySelector("assistant-message").updateComplete;
        await document.querySelector("tool-message").updateComplete;
      })()
    `);
    await page.locator('.pi-tool-card[data-tool-name="skills"] .pi-tool-card__header').click();
    const rendered = page.locator('.pi-tool-card[data-tool-name="skills"] .pi-tool-card__markdown');
    await rendered.getByRole("heading", { name: "Visible guidance" }).waitFor({ state: "visible" });

    assert.equal((await rendered.innerText()).includes("browser-sentinel"), false);
  });
});

void test("a transform bridge outage renders its setup card", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(`
      (async () => {
        const sidebar = document.querySelector("pi-sidebar");
        sidebar.agent.state.messages.push(
          {
            role: "assistant",
            content: [{ type: "toolCall", id: "bridge-tool", name: "python_transform_range", arguments: { input_range: "A1:A2", output_range: "B1:B2", code: "result = data" } }],
            api: "openai-responses",
            provider: "openai",
            model: "browser-fixture",
            usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
            stopReason: "toolUse",
            timestamp: 1,
          },
          {
            role: "toolResult",
            toolCallId: "bridge-tool",
            toolName: "python_transform_range",
            content: [{ type: "text", text: "Terminal bridge unavailable" }],
            details: {
              kind: "python_transform_range",
              blocked: false,
              bridgeUrl: "https://localhost:4441",
              error: "Terminal access is not available because the bridge is unreachable.",
              gateReason: "bridge_unreachable",
              skillHint: "python-bridge",
            },
            isError: true,
            timestamp: 2,
          },
        );
        sidebar.syncFromAgent();
        await sidebar.updateComplete;
        await document.querySelector("message-list").updateComplete;
        await document.querySelector("assistant-message").updateComplete;
        await document.querySelector("tool-message").updateComplete;
      })()
    `);

    const setupCard = page.locator(".pi-bridge-setup");
    await setupCard.waitFor({ state: "visible" });
    assert.match(await setupCard.innerText(), /python transform.*unavailable/i);
  });
});

void test("direct JavaScript tool cards identify the host runtime being used", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(`
      (async () => {
        const sidebar = document.querySelector("pi-sidebar");
        sidebar.agent.state.messages.push(
          {
            role: "assistant",
            content: [
              { type: "toolCall", id: "office-js", name: "execute_office_js", arguments: { explanation: "Inspect workbook metadata", code: "return {};" } },
              { type: "toolCall", id: "wps-js", name: "execute_wps_js", arguments: { explanation: "Inspect workbook metadata", code: "return {};" } },
            ],
            api: "openai-responses",
            provider: "openai",
            model: "browser-fixture",
            usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
            stopReason: "toolUse",
            timestamp: 1,
          },
          { role: "toolResult", toolCallId: "office-js", toolName: "execute_office_js", content: [], isError: false, timestamp: 2 },
          { role: "toolResult", toolCallId: "wps-js", toolName: "execute_wps_js", content: [], isError: false, timestamp: 3 },
        );
        sidebar.syncFromAgent();
        await sidebar.updateComplete;
        await document.querySelector("message-list").updateComplete;
        for (const element of document.querySelectorAll("assistant-message, tool-message")) {
          await element.updateComplete;
        }
      })()
    `);

    const titles = await page.locator(".pi-tool-card__title").allTextContents();
    assert.deepEqual(titles, [
      "Run Office.js Inspect workbook metadata",
      "Run WPS JSAPI Inspect workbook metadata",
    ]);
  });
});

void test("an expired login whose refresh fails asks the user to sign in again and sends no stale token", async () => {
  await withTaskpane(async (page) => {
    const sentToAnthropic: string[] = [];
    await page.route("**/v1/oauth/token", (route) => route.fulfill({
      status: 400,
      contentType: "application/json",
      headers: { "access-control-allow-origin": "*" },
      body: JSON.stringify({ error: "invalid_grant" }),
    }));
    await page.route("**/v1/messages**", (route) => {
      sentToAnthropic.push(route.request().headers()["authorization"] ?? route.request().headers()["x-api-key"] ?? "");
      return route.fulfill({ status: 401, headers: { "access-control-allow-origin": "*" }, body: "{}" });
    });
    // A Claude login stored by an earlier session; its token expired a minute ago.
    await page.evaluate(async () => {
      // Resolves from the page (/src/taskpane.html) to the module the app itself loaded.
      const storage = await import("../../src/storage/local/app-storage.ts");
      const appStorage = storage.getAppStorage();
      await appStorage.settings.set("oauth.anthropic", {
        access: "stale-access", refresh: "dead-refresh", expires: Date.now() - 60_000,
      });
      await appStorage.providerKeys.set("anthropic", "stale-access");
    });
    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    const welcome = page.locator("#pi-welcome-login-overlay");
    if (await welcome.count()) {
      await welcome.dispatchEvent("click");
      await welcome.waitFor({ state: "detached" });
    }

    await enterCommand(page, "/model");
    const picker = page.locator("#pi-model-selector-overlay");
    await picker.locator(".pi-model-selector-search").fill("claude");
    await picker.locator(".pi-model-selector-item")
      .filter({ has: page.locator(".pi-model-selector-item-provider", { hasText: /^anthropic$/u }) })
      .first().click();
    await picker.waitFor({ state: "detached" });

    await enterCommand(page, "Hello");
    await page.getByText("Your Anthropic login has expired and could not be renewed. Sign in again in Settings.", { exact: true })
      .waitFor({ state: "visible", timeout: 20_000 });
    assert.deepEqual(sentToAnthropic, [], "the expired token was sent to Anthropic");
  });
});

void test("GitHub Copilot login reaches its device code, and closing the code dialog stops the login", async () => {
  await withTaskpane(async (page, context) => {
    let tokenPolls = 0;
    // In development GitHub's login endpoints go through Vite's /oauth-proxy/github route.
    await context.route("**/oauth-proxy/github/login/device/code", (route) => route.fulfill({
      contentType: "application/json",
      body: JSON.stringify({
        device_code: "device-1",
        user_code: "ABCD-1234",
        verification_uri: "https://github.com/login/device",
        interval: 1,
        expires_in: 900,
      }),
    }));
    await context.route("**/oauth-proxy/github/login/oauth/access_token", (route) => {
      tokenPolls += 1;
      return route.fulfill({ contentType: "application/json", body: JSON.stringify({ error: "authorization_pending" }) });
    });

    await enterCommand(page, "/login");
    await page.getByText("GitHub Copilot", { exact: true }).first().click();
    const firstPoll = page.waitForResponse("**/oauth-proxy/github/login/oauth/access_token");
    await page.getByRole("button", { name: "Login with GitHub Copilot" }).click();
    // GitHub Enterprise domain: blank means github.com.
    const prompt = page.locator("#pi-prompt-overlay");
    await prompt.locator("input").waitFor({ state: "visible" });
    await prompt.getByRole("button", { name: /continue/i }).click();

    await prompt.locator("code").filter({ hasText: "ABCD-1234" }).waitFor({ state: "visible" });
    await firstPoll;

    await page.keyboard.press("Escape");
    await prompt.waitFor({ state: "detached" });
    const pollsAtClose = tokenPolls;
    await page.waitForTimeout(3_000);
    assert.equal(tokenPolls, pollsAtClose, "the login kept polling after its dialog closed");
    await page.getByRole("button", { name: "Login with GitHub Copilot" }).waitFor({ state: "visible" });
  });
});
