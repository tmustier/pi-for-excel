import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer({ token: "extension-widgets-browser-token" });
});

after(async () => {
  await env.close();
});

async function openTaskpaneWithExtension(code: string): Promise<{
  page: Page;
  finish(): Promise<void>;
}> {
  let commandSent = false;
  const opened = await openTaskpane(env, {
    clientId: "extension-widget-test",
    prepareContext: async (context) => {
      await context.addInitScript(() => {
        if (window === window.top) {
          localStorage.setItem("pi.experimental.extensionWidgetV2", "1");
        }
      });
    },
    bridge: async (url, route) => {
      if (url.pathname !== "/client/poll" || commandSent) return false;
      commandSent = true;
      await route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({
          id: "install-widget-extension",
          type: "extensionInstallCode",
          payload: { name: "Widget browser extension", code },
        }),
      });
      return true;
    },
  });
  return { page: opened.page, finish: opened.finish };
}

void test("an installed extension's widget bounds are clamped and remain valid", async () => {
  const { page, finish } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const el = document.createElement("div");
      el.textContent = "Bounded widget content";
      api.widget.upsert({
        id: "bounded",
        el,
        title: "Bounded widget",
        minHeightPx: 700,
        maxHeightPx: 12
      });
    }
  `);

  try {
    const body = page.locator("[data-widget-id='bounded'] .pi-ext-widget-body");
    await body.waitFor({ state: "visible", timeout: 10_000 });
    assert.equal(await body.evaluate((element) => element.style.minHeight), "640px");
    assert.equal(await body.evaluate((element) => element.style.maxHeight), "640px");
  } finally {
    await finish();
  }
});

void test("an installed extension preserves widget bounds when an upsert omits them", async () => {
  const { page, finish } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const first = document.createElement("div");
      first.textContent = "First bounded content";
      api.widget.upsert({ id: "preserved", el: first, minHeightPx: 180, maxHeightPx: 420 });

      const replacement = document.createElement("div");
      replacement.textContent = "Replacement bounded content";
      api.widget.upsert({ id: "preserved", el: replacement });
    }
  `);

  try {
    const body = page.locator("[data-widget-id='preserved'] .pi-ext-widget-body");
    await body.waitFor({ state: "visible", timeout: 10_000 });
    assert.deepEqual(
      await body.evaluate((element) => ({ min: element.style.minHeight, max: element.style.maxHeight })),
      { min: "180px", max: "420px" },
    );
  } finally {
    await finish();
  }
});

void test("an installed extension can clear existing widget bounds", async () => {
  const { page, finish } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const first = document.createElement("div");
      first.textContent = "Initially bounded content";
      api.widget.upsert({ id: "cleared", el: first, minHeightPx: 180, maxHeightPx: 420 });

      const replacement = document.createElement("div");
      replacement.textContent = "Unbounded replacement content";
      api.widget.upsert({ id: "cleared", el: replacement, minHeightPx: null, maxHeightPx: null });
    }
  `);

  try {
    const body = page.locator("[data-widget-id='cleared'] .pi-ext-widget-body");
    await body.waitFor({ state: "visible", timeout: 10_000 });
    assert.deepEqual(
      await body.evaluate((element) => ({ min: element.style.minHeight, max: element.style.maxHeight })),
      { min: "", max: "" },
    );
  } finally {
    await finish();
  }
});

void test("an installed extension cannot leave a non-collapsible widget collapsed", async () => {
  const { page, finish } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const el = document.createElement("div");
      el.textContent = "Always visible widget content";
      api.widget.upsert({
        id: "always-visible",
        el,
        title: "Always visible widget",
        collapsible: false,
        collapsed: true
      });
    }
  `);

  try {
    const card = page.locator("[data-widget-id='always-visible']");
    const body = card.locator(".pi-ext-widget-body");
    await body.waitFor({ state: "visible", timeout: 10_000 });
    assert.equal(await card.getByRole("button").count(), 0);
    assert.equal(await body.isVisible(), true);
  } finally {
    await finish();
  }
});

void test("an installed extension preserves collapsed state when an upsert omits it", async () => {
  const { page, finish } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const first = document.createElement("div");
      first.textContent = "Initially collapsed content";
      api.widget.upsert({ id: "collapsed", el: first, title: "Collapsed widget", collapsible: true, collapsed: true });

      const replacement = document.createElement("div");
      replacement.textContent = "Replacement collapsed content";
      api.widget.upsert({ id: "collapsed", el: replacement });
    }
  `);

  try {
    const card = page.locator("[data-widget-id='collapsed']");
    await card.waitFor({ state: "visible", timeout: 10_000 });
    assert.equal(await card.locator(".pi-ext-widget-body").isVisible(), false);
    assert.equal(await card.getByRole("button").getAttribute("aria-expanded"), "false");
  } finally {
    await finish();
  }
});
