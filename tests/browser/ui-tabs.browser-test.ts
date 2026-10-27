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

async function renameTabInDialog(page: Page, title: string): Promise<void> {
  const dialog = page.locator("#pi-text-input-dialog-overlay");
  const input = dialog.getByRole("textbox");
  await input.waitFor({ state: "visible" });
  await input.fill(title);
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
    await renameTabInDialog(page, "Revenue Model");
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
    await page.keyboard.press("Control+Alt+Shift+T");
    await page.locator(".pi-session-tab.is-active .pi-session-tab__title").filter({ hasText: "Revenue Model" })
      .waitFor({ state: "visible" });
  });
});

void test("Ctrl+Alt arrows switch tabs from tab or composer focus, and tab shortcuts ignore an open dialog", async () => {
  await withTaskpane(async (page) => {
    await page.getByRole("button", { name: "New tab" }).click();
    const active = page.locator(".pi-session-tab.is-active .pi-session-tab__title");
    await active.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });

    // A new tab moves focus to the composer on the next frame; let that land
    // before focusing the tab, or it takes focus back after the key press.
    await page.waitForFunction(() => document.activeElement?.closest("pi-input") !== null);
    await page.locator(".pi-session-tab.is-active .pi-session-tab__main").focus();
    await page.waitForFunction(() => document.activeElement?.closest(".pi-session-tab") !== null);
    await page.keyboard.press("Control+Alt+ArrowRight");
    await active.filter({ hasText: "Chat 1" }).waitFor({ state: "visible" });
    await page.locator("pi-input textarea").focus();
    await page.keyboard.press("Control+Alt+ArrowLeft");
    await active.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });

    await page.locator("pi-input textarea").fill("/settings");
    await page.locator("pi-input textarea").press("Enter");
    await page.locator("#pi-settings-overlay").waitFor({ state: "visible" });
    await page.keyboard.press("Control+Alt+n");
    await page.keyboard.press("Control+Alt+w");
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
        await renameTabInDialog(page, "  Forecast  ");
      }

      await page.locator(".pi-session-tab.is-active .pi-session-tab__title").filter({ hasText: row.expected }).waitFor({ state: "visible" });
      assert.equal(await page.locator(".pi-session-tab.is-active .pi-session-tab__title").innerText(), row.expected);
    }
  });
});

void test("Ctrl/Cmd+Z never reopens tabs; common tab chords work in the composer", async () => {
  await withTaskpane(async (page) => {
    const tabs = page.locator(".pi-session-tab");
    const composer = page.locator("pi-input textarea");
    await composer.focus();
    await page.keyboard.press("Control+Alt+n");
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });

    await tabs.filter({ hasText: "Chat 2" }).getByRole("button", { name: "Close tab" }).click();
    await page.locator("#pi-toast").getByText("Closed Chat 2").waitFor({ state: "visible" });
    await composer.click();
    await page.keyboard.press("Control+z");
    await page.keyboard.press("Meta+z");
    assert.equal(await tabs.count(), 1);
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

    await page.keyboard.press("Control+Alt+Shift+T");
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "visible" });
    await composer.focus();
    await page.keyboard.press("Control+Alt+w");
    await tabs.filter({ hasText: "Chat 2" }).waitFor({ state: "detached" });
  });
});

void test("clicking anywhere on an inactive tab's title switches to it", async () => {
  await withTaskpane(async (page) => {
    await page.getByRole("button", { name: "New tab" }).click();
    await page.locator(".pi-session-tab").nth(1).and(page.locator(".is-active")).waitFor();
    const title = page.locator(".pi-session-tab").first().locator(".pi-session-tab__title");
    const box = await title.boundingBox();
    assert.ok(box);
    // The right end of a short title is where an overlaid × would sit.
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.click(box.x + box.width - 2, box.y + box.height / 2);
    await page.locator(".pi-session-tab").first().and(page.locator(".is-active")).waitFor({ timeout: 3_000 });
    assert.equal(await page.locator(".pi-session-tab").count(), 2);
  });
});

void test("the active last tab stays fully visible when a turn widens it", async () => {
  await withTaskpane(async (page) => {
    await page.setViewportSize({ width: 349, height: 580 });
    for (let i = 1; i < 12; i++) {
      await page.getByRole("button", { name: "New tab" }).click();
    }
    await page.locator(".pi-session-tab.is-active").filter({ hasText: "Chat 12" }).waitFor();
    await page.evaluate(async () => {
      const { PiSidebar } = await import("/src/ui/pi-sidebar.ts") as typeof import("../../src/ui/pi-sidebar.ts");
      const sidebar = document.querySelector("pi-sidebar");
      if (!(sidebar instanceof PiSidebar) || !sidebar.agent) throw new Error("sidebar missing");
      sidebar.agent.getApiKey = () => Promise.resolve("test-key");
      sidebar.agent.streamFunction = (_model, _context, options) => new Promise<never>((_resolve, reject) => {
        options?.signal?.addEventListener("abort", () => reject(new Error("aborted")));
      });
    });
    const welcome = page.locator("#pi-welcome-login-overlay");
    if (await welcome.count()) await welcome.dispatchEvent("click");
    const composer = page.locator("pi-input textarea");
    await composer.fill("Keep working");
    await composer.press("Enter");
    await page.locator(".pi-session-tab.is-active .pi-session-tab__busy").waitFor();
    await page.waitForTimeout(200);
    const close = await page.locator(".pi-session-tab.is-active .pi-session-tab__close").boundingBox();
    const strip = await page.locator(".pi-session-tabs__scroller").boundingBox();
    assert.ok(close && strip);
    assert.ok(close.x >= strip.x && close.x + close.width <= strip.x + strip.width + 1,
      "the busy tab's close button must stay inside the visible strip");
    const scroller = page.locator(".pi-session-tabs__scroller");
    const previousScroll = await scroller.evaluate((element) => element.scrollLeft);
    await page.getByRole("button", { name: "Scroll tabs left" }).click();
    await page.waitForTimeout(400);
    assert.ok(await scroller.evaluate((element) => element.scrollLeft) < previousScroll - 100,
      "the active tab must not pull the strip back after a manual scroll");
    await page.locator(".pi-input-btn--abort").click();
  });
});
