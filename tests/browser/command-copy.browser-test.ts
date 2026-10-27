import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { openBrowserPage, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

void test("the proxy warning copies the current helper command without an npm install prompt", async () => {
  const opened = await openBrowserPage(env, { path: "src/ui-gallery.html" });
  await opened.context.grantPermissions(["clipboard-read", "clipboard-write"], { origin: new URL(env.baseUrl).origin });
  const { page } = opened;
  try {
    await page.evaluate(async () => {
      const { createProxyBannerController } = await import("/src/ui/proxy-banner.ts") as typeof import("../../src/ui/proxy-banner.ts");
      const notices = document.createElement("div");
      notices.id = "pi-notices";
      document.body.prepend(notices);
      await createProxyBannerController({ isProxyEnabled: () => Promise.resolve(true) }).update("not-detected");
    });
    const banner = page.locator("#pi-notices");
    await banner.getByRole("button", { name: "Show how to start it" }).click();
    assert.equal(await banner.locator("code").innerText(), "npx -y pi-for-excel-proxy@latest");
    await banner.getByRole("button", { name: "Copy command" }).click();
    assert.equal(await page.evaluate(() => navigator.clipboard.readText()), "npx -y pi-for-excel-proxy@latest");
  } finally {
    await opened.finish();
  }
});

void test("a command row copies, translates its accessible state, and resets it", async () => {
  const opened = await openBrowserPage(env, { path: "src/ui-gallery.html" });
  await opened.context.grantPermissions(["clipboard-read", "clipboard-write"], { origin: new URL(env.baseUrl).origin });
  const { page } = opened;

  try {
    await page.evaluate(async () => {
      // Vite resolves these browser modules; the assertions describe their checked source contracts.
      const language = await import("/src/language/index.ts") as typeof import("../../src/language/index.ts");
      const commandCopy = await import("/src/ui/command-copy.ts") as typeof import("../../src/ui/command-copy.ts");
      language.initLanguage("zh-CN");
      document.body.append(commandCopy.createCopyableCommand("npx pi-for-excel-proxy --browser-contract"));
    });

    const row = page.locator(".pi-command-copy").filter({ hasText: "--browser-contract" });
    const button = row.getByRole("button", { name: "复制命令" });
    await button.click();

    await page.getByRole("button", { name: "已复制" }).waitFor({ state: "visible" });
    assert.equal(await page.evaluate(() => navigator.clipboard.readText()), "npx pi-for-excel-proxy --browser-contract");
    await page.getByRole("button", { name: "复制命令" }).waitFor({ state: "visible", timeout: 2_500 });
  } finally {
    await opened.finish();
  }
});
