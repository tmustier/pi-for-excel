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

void test("theme follows Office by default, permits persistent overrides, and keeps native controls dark", async () => {
  await withTaskpane(async (page) => {
    await page.emulateMedia({ colorScheme: "dark" });
    assert.equal(await page.locator("html").evaluate((root) => root.classList.contains("dark")), false);
    await enterCommand(page, "/settings");
    const theme = page.getByRole("combobox", { name: "Theme", exact: true });
    assert.equal(await theme.inputValue(), "office");
    await theme.selectOption("dark");
    await page.waitForFunction(() => document.documentElement.classList.contains("dark"));
    assert.equal(await page.locator("html").evaluate((root) => getComputedStyle(root).colorScheme), "dark");
    await page.reload();
    await page.locator("pi-input textarea").waitFor();
    assert.equal(await page.locator("html").evaluate((root) => root.classList.contains("dark")), true);
    await page.locator("#pi-welcome-login-overlay").dispatchEvent("click");
    await enterCommand(page, "/settings");
    await theme.selectOption("light");
    await page.waitForFunction(() => !document.documentElement.classList.contains("dark"));
    await theme.selectOption("office");
    await page.evaluate(() => Reflect.set(window, "Office", {
      context: { officeTheme: { isDarkTheme: true }, host: "Excel" },
    }));
    await page.waitForFunction(() => document.documentElement.classList.contains("dark"));
    await theme.selectOption("light");
    await page.waitForFunction(() => !document.documentElement.classList.contains("dark"));
  });
});
