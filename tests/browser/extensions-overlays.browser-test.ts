import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer({ token: "extension-overlays-browser-token" });
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
    clientId: "extension-overlay-test",
    bridge: async (url, route) => {
      if (url.pathname !== "/client/poll" || commandSent) return false;
      commandSent = true;
      await route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({
          id: "install-overlay-extension",
          type: "extensionInstallCode",
          payload: { name: "Overlay browser extension", code },
        }),
      });
      return true;
    },
    welcomeClickForce: true,
  });
  return { page: opened.page, finish: opened.finish };
}

void test("an installed extension can replace and dismiss its visible overlay", async () => {
  const { page, finish } = await openTaskpaneWithExtension(`
    export function activate(api) {
      const first = document.createElement("section");
      first.textContent = "First extension overlay";
      api.overlay.show(first);

      setTimeout(() => {
        const second = document.createElement("section");
        second.textContent = "Replacement extension overlay";
        api.overlay.show(second);
      }, 1000);

      setTimeout(() => api.overlay.dismiss(), 2000);
    }
  `);

  try {
    const overlay = page.locator("#pi-ext-overlay");
    await overlay.getByText("First extension overlay").waitFor({ state: "visible", timeout: 10_000 });

    await overlay.getByText("Replacement extension overlay").waitFor({ state: "visible", timeout: 5_000 });
    assert.equal(await overlay.getByText("First extension overlay").count(), 0);

    await overlay.waitFor({ state: "detached", timeout: 5_000 });
  } finally {
    await finish();
  }
});
