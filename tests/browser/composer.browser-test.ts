import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { enterInput, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

async function withTaskpane(run: (page: Page) => Promise<void>, viewportHeight = 720): Promise<void> {
  const opened = await openTaskpane(env, {
    viewport: { width: 420, height: viewportHeight },
    clientId: "composer-browser-client",
  });
  try {
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

void test("Alt+Up recalls queued actions into the composer in order", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(`
      (async () => {
        const { commandRegistry } = await import("/src/commands/types.ts");
        commandRegistry.register({
          name: "compact",
          description: "Keep the action queue busy",
          source: "builtin",
          execute: () => new Promise(() => {}),
        });
      })()
    `);

    await enterInput(page, "/compact");
    await enterInput(page, "first queued action");
    await page.locator("#pi-queue-display").getByText("first queued action", { exact: true }).waitFor({ state: "visible" });
    await enterInput(page, "second queued action");
    await page.locator("#pi-queue-display").getByText("second queued action", { exact: true }).waitFor({ state: "visible" });

    const input = page.locator("pi-input textarea");
    await input.press("Alt+ArrowUp");

    assert.deepEqual(
      {
        composer: await input.inputValue(),
        queuedItems: await page.locator("#pi-queue-display .pi-queue__item").count(),
      },
      {
        composer: "first queued action\n\nsecond queued action",
        queuedItems: 0,
      },
    );
  });
});

void test("the composer trims a prompt and does not submit whitespace", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await page.locator("pi-input").evaluate((composer) => {
      composer.addEventListener("pi-send", (event) => {
        // The component owns this event contract; the assertion only exposes its browser payload for observation.
        const sendEvent = event as CustomEvent<{ text?: string }>;
        if (typeof sendEvent.detail?.text === "string") {
          document.body.dataset.lastComposerSend = sendEvent.detail.text;
        }
      });
    });
    await input.fill("   \n\t  ");
    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-last-composer-send"), null);

    await input.fill("  build a TSLA P&L  ");
    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-last-composer-send"), "build a TSLA P&L");
  });
});

void test("Shift+Enter edits while Enter submits a real prompt", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await page.locator("pi-input").evaluate((composer) => {
      composer.addEventListener("pi-send", () => { document.body.dataset.composerSent = "yes"; });
    });
    await input.fill("write A1");
    await input.press("Shift+Enter");
    assert.equal(await page.locator("body").getAttribute("data-composer-sent"), null);
    assert.equal(await input.inputValue(), "write A1\n");

    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-composer-sent"), "yes");
  });
});

void test("composer auto-grow honors its CSS cap", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.fill(Array.from({ length: 50 }, () => "row").join("\n"));
    assert.equal(await input.evaluate((element) => element.style.height), "288px");
  }, 720);
});

void test("composer auto-grow uses the smaller short-pane cap", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.evaluate((element) => { element.style.maxHeight = "none"; });
    await input.fill(Array.from({ length: 50 }, () => "row").join("\n"));
    assert.equal(await input.evaluate((element) => element.style.height), "140px");
  }, 500);
});
