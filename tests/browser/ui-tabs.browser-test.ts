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

void test("renamed and closed tabs restore with the active tab after reload", async () => {
  await withTaskpane(async (page) => {
    const newTab = page.getByRole("button", { name: "New tab" });
    await newTab.click();
    const tabs = page.locator(".pi-session-tab");
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });
    await newTab.click();
    await tabs.filter({ hasText: "Chat 3" }).waitFor({ state: "visible" });

    page.once("dialog", (dialog) => dialog.accept("Revenue Model"));
    await tabs.filter({ hasText: "Chat 2" }).locator(".pi-session-tab__main").dblclick();
    await tabs.filter({ hasText: "Revenue Model" }).waitFor({ state: "visible" });

    await tabs.filter({ hasText: "Revenue Model" }).getByRole("button", { name: "Close tab" }).click();
    await tabs.filter({ hasText: "Chat 3" }).locator(".pi-session-tab__main").click();
    await page.waitForTimeout(300);

    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });

    assert.deepEqual(await page.locator(".pi-session-tab__title").allTextContents(), ["Chat 1", "Chat 2"]);
    assert.equal(await page.locator(".pi-session-tab.is-active .pi-session-tab__title").innerText(), "Chat 2");
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
        page.once("dialog", (dialog) => dialog.accept("  Forecast  "));
        await page.locator(".pi-session-tab.is-active .pi-session-tab__main").dblclick();
      }

      await page.locator(".pi-session-tab.is-active .pi-session-tab__title").filter({ hasText: row.expected }).waitFor({ state: "visible" });
      assert.equal(await page.locator(".pi-session-tab.is-active .pi-session-tab__title").innerText(), row.expected);
    }
  });
});
