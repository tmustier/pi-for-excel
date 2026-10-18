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
  const opened = await openTaskpane(env, { clientId: "gateway-browser" });
  try {
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

/** Names exposed to assistive technology. */
async function exposedAccessibleText(page: Page): Promise<string[]> {
  const cdp = await page.context().newCDPSession(page);
  try {
    const { nodes } = await cdp.send("Accessibility.getFullAXTree");
    return nodes.flatMap((node) => {
      const name: unknown = node.name?.value; // CDP types AXValue.value as any.
      return !node.ignored && typeof name === "string" && name.length > 0 ? [name] : [];
    });
  } finally {
    await cdp.detach();
  }
}

void test("custom gateway deletion persists only after overlay confirmation", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.fill("/settings");
    await input.press("Enter");
    const settings = page.locator("#pi-settings-overlay");
    await settings.getByRole("button", { name: "Custom gateway" }).click();

    await page.getByPlaceholder("Gateway name (optional)").fill("Browser gateway");
    await page.getByPlaceholder("https://your-gateway.example.com/v1").fill("https://gateway.example.com/v1");
    await page.getByPlaceholder("model-id").fill("browser-model");
    await page.getByRole("button", { name: "Save gateway" }).click();
    const gateway = settings.getByText("Browser gateway", { exact: true });
    await gateway.waitFor({ state: "visible" });

    await settings.getByRole("button", { name: "Delete" }).click();
    const confirmation = page.locator("#pi-confirm-dialog-overlay");
    await confirmation.waitFor({ state: "visible" });
    // Assistive technology sees only the topmost dialog, not the settings under it.
    const exposedText = await exposedAccessibleText(page);
    assert.equal(exposedText.includes("Delete custom gateway?"), true);
    assert.equal(exposedText.includes("Configured gateways"), false);
    await confirmation.locator(".pi-overlay-btn--ghost").click();
    await confirmation.waitFor({ state: "detached" });
    await page.waitForTimeout(250);
    assert.equal(await gateway.isVisible(), true);
    assert.equal((await exposedAccessibleText(page)).includes("Configured gateways"), true);

    await settings.getByRole("button", { name: "Delete" }).click();
    await confirmation.getByRole("button", { name: "Delete" }).click();
    await gateway.waitFor({ state: "detached" });
    assert.equal(await settings.getByText("No custom gateways configured yet.", { exact: true }).isVisible(), true);
  });
});
