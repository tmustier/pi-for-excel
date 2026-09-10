import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { enterCommand, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, {
    viewport: { width: 420, height: 720 },
    clientId: "ui-settings-browser",
  });
  try {
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

void test("Settings navigation keeps an edited page open when closing is declined", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/settings");
    const settings = page.locator("#pi-settings-overlay");
    await settings.getByRole("button", { name: /Rules & conventions/ }).click();

    const rules = page.getByPlaceholder(/Your preferences and habits/);
    await rules.fill("Always show units");
    await settings.getByRole("button", { name: "Close Settings" }).click();

    const confirmation = page.locator("#pi-confirm-dialog-overlay");
    await confirmation.getByRole("heading", { name: "Discard changes?" }).waitFor({ state: "visible" });
    await confirmation.getByRole("button", { name: "Keep editing" }).filter({ hasText: "Keep editing" }).click();

    assert.equal(await rules.inputValue(), "Always show units");
  });
});
