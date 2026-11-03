import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

void test("command menu active descendant follows keyboard and pointer selection", async () => {
  const { page, finish } = await openTaskpane(env);
  try {
    const input = page.locator("pi-input textarea");
    await input.fill("/");
    const options = page.locator("#pi-command-menu [role=option]");
    await options.nth(2).waitFor({ state: "visible" });
    await input.press("ArrowDown");
    assert.equal(await input.getAttribute("aria-activedescendant"), await options.nth(1).getAttribute("id"), "keyboard selection must update the active descendant");
    await options.nth(2).hover();
    assert.equal(await input.getAttribute("aria-activedescendant"), await options.nth(2).getAttribute("id"), "pointer selection must update the active descendant");
    await input.press("Escape");
    assert.equal(await input.getAttribute("aria-activedescendant"), null);
  } finally { await finish(); }
});
