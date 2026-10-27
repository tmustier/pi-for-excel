import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

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
    clientId: "ui-tabs-browser",
  });
  try {
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

async function renameInDialog(page: Page, title: string): Promise<void> {
  const dialog = page.locator("#pi-text-input-dialog-overlay");
  await dialog.locator("input").fill(title);
  await dialog.getByRole("button", { name: "Rename" }).click();
  await dialog.waitFor({ state: "detached" });
}

void test("renamed and closed tabs restore with the active tab after reload", async () => {
  await withTaskpane(async (page) => {
    const newTab = page.getByRole("button", { name: "New tab" });
    await newTab.click();
    const tabs = page.locator(".pi-session-tab");
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });
    await newTab.click();
    await tabs.filter({ hasText: "Chat 3" }).waitFor({ state: "visible" });

    await tabs.filter({ hasText: "Chat 2" }).locator(".pi-session-tab__main").dblclick();
    await renameInDialog(page, "Revenue Model");
    await tabs.filter({ hasText: "Revenue Model" }).waitFor({ state: "visible" });

    await tabs.filter({ hasText: "Revenue Model" }).getByRole("button", { name: "Close tab" }).click();
    await tabs.filter({ hasText: "Chat 3" }).locator(".pi-session-tab__main").click();
    await page.waitForTimeout(300);

    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });

    // Tabs keep their numbers across a reload: the closed "Chat 2" is not reused.
    assert.deepEqual(await page.locator(".pi-session-tab__title").allTextContents(), ["Chat 1", "Chat 3"]);
    assert.equal(await page.locator(".pi-session-tab.is-active .pi-session-tab__title").innerText(), "Chat 3");

    // The closed tab is still in recently closed after the reload, and comes back with its name.
    const welcome = page.locator("#pi-welcome-login-overlay");
    await welcome.dispatchEvent("click");
    await welcome.waitFor({ state: "detached" });
    await page.locator("pi-input textarea").click();
    await page.keyboard.press("Control+Shift+T");
    await page.locator(".pi-session-tab.is-active .pi-session-tab__title").filter({ hasText: "Revenue Model" })
      .waitFor({ state: "visible" });
  });
});

void test("Ctrl+PageDown switches tabs while a tab has focus, and tab shortcuts ignore an open dialog", async () => {
  await withTaskpane(async (page) => {
    await page.getByRole("button", { name: "New tab" }).click();
    const active = page.locator(".pi-session-tab.is-active .pi-session-tab__title");
    await active.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });

    // A new tab moves focus to the composer on the next frame; let that land
    // before focusing the tab, or it takes focus back after the key press.
    await page.waitForFunction(() => document.activeElement?.closest("pi-input") !== null);
    await page.locator(".pi-session-tab.is-active .pi-session-tab__main").focus();
    await page.waitForFunction(() => document.activeElement?.closest(".pi-session-tab") !== null);
    await page.keyboard.press("Control+PageDown");
    await active.filter({ hasText: "Chat 1" }).waitFor({ state: "visible" });
    await page.keyboard.press("Control+PageUp");
    await active.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });

    await page.locator("pi-input textarea").fill("/settings");
    await page.locator("pi-input textarea").press("Enter");
    await page.locator("#pi-settings-overlay").waitFor({ state: "visible" });
    await page.keyboard.press("Control+t");
    await page.keyboard.press("Control+w");
    await page.waitForTimeout(300);
    assert.equal(await page.locator(".pi-session-tab").count(), 2);
    assert.equal(await active.innerText(), "Chat 2");
  });
});

void test("visible tab titles follow default numbering and explicit renames", async () => {
  await withTaskpane(async (page) => {
    const cases = [
      { action: "initial", expected: "Chat 1" },
      { action: "new", expected: "Chat 2" },
      { action: "rename", expected: "Forecast" },
    ] as const;

    for (const row of cases) {
      if (row.action === "new") {
        await page.getByRole("button", { name: "New tab" }).click();
      } else if (row.action === "rename") {
        await page.locator(".pi-session-tab.is-active .pi-session-tab__main").dblclick();
        await renameInDialog(page, "  Forecast  ");
      }

      await page.locator(".pi-session-tab.is-active .pi-session-tab__title").filter({ hasText: row.expected }).waitFor({ state: "visible" });
      assert.equal(await page.locator(".pi-session-tab.is-active .pi-session-tab__title").innerText(), row.expected);
    }
  });
});

void test("Ctrl+Z reopens a closed tab only while its Undo toast is showing", async () => {
  await withTaskpane(async (page) => {
    const tabs = page.locator(".pi-session-tab");
    const newTab = page.getByRole("button", { name: "New tab" });
    const composer = page.locator("pi-input textarea");
    await newTab.click();
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });

    await tabs.filter({ hasText: "Chat 2" }).getByRole("button", { name: "Close tab" }).click();
    await page.locator("#pi-toast").getByText("Closed Chat 2").waitFor({ state: "visible" });
    await composer.click();
    await page.keyboard.press("Control+z");
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });

    await tabs.filter({ hasText: "Chat 2" }).getByRole("button", { name: "Close tab" }).click();
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "detached" });
    // After a reload the toast is gone but the tab is still in recently closed.
    await page.reload();
    await composer.waitFor({ state: "visible", timeout: 20_000 });
    const welcome = page.locator("#pi-welcome-login-overlay");
    await welcome.dispatchEvent("click");
    await welcome.waitFor({ state: "detached" });
    await composer.click();
    await page.keyboard.press("Control+z");
    await page.waitForTimeout(500);
    assert.deepEqual(await page.locator(".pi-session-tab__title").allTextContents(), ["Chat 1"]);

    await page.keyboard.press("Control+Shift+T");
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });
  });
});
