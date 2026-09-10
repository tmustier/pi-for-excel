import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { chromium, type Browser } from "playwright";
import { createServer, type ViteDevServer } from "vite";

let browser: Browser;
let server: ViteDevServer;
let baseUrl: string;

before(async () => {
  server = await createServer({
    configFile: false,
    root: process.cwd(),
    server: { host: "127.0.0.1", port: 0, strictPort: false },
  });
  await server.listen();
  const localUrl = server.resolvedUrls?.local[0];
  if (!localUrl) throw new Error("Vite did not expose a local URL");
  baseUrl = localUrl;
  browser = await chromium.launch({ headless: true });
});

after(async () => {
  await browser?.close();
  await server?.close();
});

void test("a command row copies, translates its accessible state, and resets it", async () => {
  const context = await browser.newContext();
  await context.grantPermissions(["clipboard-read", "clipboard-write"], { origin: new URL(baseUrl).origin });
  const page = await context.newPage();

  try {
    await page.goto(`${baseUrl}src/ui-gallery.html`);
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
    await context.close();
  }
});
