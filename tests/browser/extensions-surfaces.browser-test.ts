import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer({ token: "extensions-surfaces-token" });
});

after(async () => {
  await env.close();
});

void test("Plugins shows the installed extensions and installation sections", async () => {
  const opened = await openTaskpane(env, { clientId: "extensions-hub" });
  const { page } = opened;

  try {
    const input = page.locator("pi-input textarea");
    await input.fill("/plugins");
    await input.press("Enter");

    const plugins = page.locator("#pi-settings-overlay");
    await plugins.getByRole("heading", { name: "Plugins" }).waitFor({ state: "visible" });
    const sectionLabels = await plugins.locator(".pi-section-header__label").allTextContents();
    assert.deepEqual(sectionLabels.filter((label) => label === "Installed" || label === "Install"), [
      "Installed",
      "Install",
    ]);
  } finally {
    await opened.finish();
  }
});

void test("an installed extension can remove one widget and then clear its remaining widgets", async () => {
  let commandSent = false;
  const opened = await openTaskpane(env, {
    clientId: "extensions-widgets",
    prepareContext: async (context) => {
      await context.addInitScript(() => {
        if (window === window.top) localStorage.setItem("pi.experimental.extensionWidgetV2", "1");
      });
    },
    bridge: async (url, route) => {
      if (url.pathname !== "/client/poll" || commandSent) return false;
      commandSent = true;
      await route.fulfill({
        contentType: "application/json",
        body: JSON.stringify({
          id: "install-extensions-widgets",
          type: "extensionInstallCode",
          payload: {
            name: "Widget lifecycle extension",
            code: `
              export function activate(api) {
                for (const id of ["discarded", "remaining"]) {
                  const el = document.createElement("div");
                  el.textContent = id;
                  api.widget.upsert({ id, el });
                }
                setTimeout(() => api.widget.remove("discarded"), 100);
                setTimeout(() => api.widget.clear(), 500);
              }
            `,
          },
        }),
      });
      return true;
    },
  });
  const { page } = opened;

  try {
    const remaining = page.locator("[data-widget-id='remaining']");
    await remaining.waitFor({ state: "visible", timeout: 10_000 });
    await page.locator("[data-widget-id='discarded']").waitFor({ state: "detached", timeout: 5_000 });
    await remaining.waitFor({ state: "detached", timeout: 5_000 });
    assert.equal(await page.locator(".pi-ext-widget-card").count(), 0);
  } finally {
    await opened.finish();
  }
});
