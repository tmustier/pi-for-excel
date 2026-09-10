import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer({ token: "extensions-browser-token" });
});

after(async () => {
  await env.close();
});

void test("an inline extension installed through the runtime appears enabled in Plugins", async () => {
  let commandSent = false;
  let resolveInstalled: (() => void) | undefined;
  const installed = new Promise<void>((resolve) => {
    resolveInstalled = resolve;
  });

  const opened = await openTaskpane(env, {
    clientId: "extension-hub-browser",
    bridge: async (url, route) => {
      if (url.pathname === "/client/poll" && !commandSent) {
        commandSent = true;
        await route.fulfill({
          contentType: "application/json",
          body: JSON.stringify({
            id: "install-inline-extension",
            type: "extensionInstallCode",
            payload: {
              name: "Browser audit extension",
              code: "export function activate() {}",
            },
          }),
        });
      } else if (url.pathname === "/client/result") {
        resolveInstalled?.();
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true }) });
      } else if (url.pathname === "/client/poll") {
        await new Promise((resolve) => setTimeout(resolve, 250));
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ type: "noop" }) });
      } else {
        return false;
      }
      return true;
    },
  });
  const { page } = opened;
  const input = page.locator("pi-input textarea");

  try {
    await Promise.race([
      installed,
      new Promise<void>((_resolve, reject) => {
        setTimeout(() => reject(new Error("inline extension installation did not complete")), 10_000);
      }),
    ]);

    await input.fill("/plugins");
    await input.press("Enter");

    const plugins = page.locator("#pi-settings-overlay");
    await plugins.getByRole("heading", { name: "Plugins" }).waitFor({ state: "visible", timeout: 5_000 });
    const extensionCard = plugins.locator(".pi-item-card", { hasText: "Browser audit extension" });
    await extensionCard.waitFor({ state: "visible" });
    assert.match((await extensionCard.textContent()) ?? "", /inline code \(29 chars, 1 lines\) · sandbox iframe/u);
    await assert.doesNotReject(() => extensionCard.getByRole("checkbox").waitFor({ state: "attached" }));
    assert.equal(await extensionCard.getByRole("checkbox").isChecked(), true);
  } finally {
    await opened.finish();
  }
});
