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

async function prepareTaskpanePage(context: BrowserContext): Promise<Page> {
  const page = await context.newPage();
  const pageErrors: Error[] = [];

  page.on("pageerror", (error) => {
    if (!isExpectedOfficeUnavailableError(error)) pageErrors.push(error);
  });
  await context.route("**/*", async (route) => {
    const url = new URL(route.request().url());
    if (url.hostname === "localhost" && url.port === "3157") {
      const body = url.pathname === "/client/register"
        ? { clientId: "ui-taskpane-browser" }
        : { type: "noop" };
      await route.fulfill({ contentType: "application/json", body: JSON.stringify(body) });
      return;
    }
    if (["127.0.0.1", "localhost", "[::1]"].includes(url.hostname)) {
      await route.continue();
      return;
    }
    await route.abort("blockedbyclient");
  });

  await page.goto(`${baseUrl}src/taskpane.html`);
  await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
  const welcome = page.locator("#pi-welcome-login-overlay");
  await welcome.waitFor({ state: "visible", timeout: 10_000 });
  await welcome.click({ position: { x: 2, y: 2 } });
  await welcome.waitFor({ state: "detached" });
  await page.evaluate(() => {
    window.addEventListener("beforeunload", () => {
      document.body.dataset.browserPageErrorsChecked = "true";
    });
  });
  assert.deepEqual(pageErrors.map((error) => error.message), []);
  return page;
}

async function withTaskpane(run: (page: Page, context: BrowserContext) => Promise<void>): Promise<void> {
  const context = await browser.newContext({ viewport: { width: 420, height: 720 } });
  try {
    const page = await prepareTaskpanePage(context);
    await run(page, context);
  } finally {
    await context.close();
  }
}

async function enterCommand(page: Page, command: string): Promise<void> {
  const input = page.locator("pi-input textarea");
  await input.fill(command);
  await input.press("Enter");
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

void test("renaming a tab confirms the entered visible title", async () => {
  await withTaskpane(async (page) => {
    page.once("dialog", async (dialog) => {
      assert.equal(dialog.type(), "prompt");
      await dialog.accept("Revenue Model");
    });
    await page.locator(".pi-session-tab__main").dblclick();

    await page.locator(".pi-session-tab__title").filter({ hasText: "Revenue Model" }).waitFor({ state: "visible" });
    assert.equal(await page.locator(".pi-session-tab__title").innerText(), "Revenue Model");
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
