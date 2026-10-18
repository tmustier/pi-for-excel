import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { openBrowserPage, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let server: TaskpaneServer;
before(async () => { server = await startTaskpaneServer(); });
after(async () => { await server.close(); });

void test("inspection, stale acceptance, explicit second acknowledgement and Escape use real controls", async () => {
  const opened = await openBrowserPage(server, { path: "src/ui-gallery.html", viewport: { width: 350, height: 780 } });
  const { page } = opened;
  try {
    const section = page.locator('[data-gallery="durable-conflict"]');
    const inspect = section.getByRole("button", { name: "Inspect", exact: true });
    await inspect.click();
    const dialog = page.getByRole("dialog");
    await dialog.getByText("Original evidence", { exact: true }).waitFor();
    assert.match(await dialog.textContent() ?? "", /observation-1/);
    assert.match(await dialog.textContent() ?? "", /not proof of success/);
    await dialog.getByRole("button", { name: "Acknowledge current state" }).click();
    await dialog.getByRole("alert").waitFor();
    assert.match(await section.textContent() ?? "", /1 workbook changes require review/);
    assert.match(await dialog.textContent() ?? "", /observation-2/);
    await dialog.getByRole("button", { name: "Acknowledge current state" }).waitFor();
    await dialog.getByRole("button", { name: "Acknowledge current state" }).click();
    await dialog.waitFor({ state: "detached" });
    await section.getByText("No workbook changes awaiting review.", { exact: true }).waitFor();
    assert.match(await section.textContent() ?? "", /original success is not asserted/);
    assert.equal(await section.getByRole("button", { name: "Refresh reviews" }).evaluate(node => node === document.activeElement), true);
    const pendingInspect = page.locator('[data-gallery="durable-pending"]').getByRole("button", { name: "Inspect", exact: true });
    await pendingInspect.click();
    await page.keyboard.press("Escape");
    await dialog.waitFor({ state: "detached" });
    assert.equal(await pendingInspect.evaluate(node => node === document.activeElement), true);
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  } finally { await opened.finish(); }
});

void test("inspection errors retain pending and offer no acknowledgement", async () => {
  const opened = await openBrowserPage(server, { path: "src/ui-gallery.html" });
  try {
    const section = opened.page.locator('[data-gallery="durable-error"]');
    await section.getByRole("button", { name: "Inspect", exact: true }).click();
    const dialog = opened.page.getByRole("dialog");
    await dialog.getByRole("alert").waitFor();
    await dialog.getByText("Original evidence", { exact: true }).waitFor();
    assert.equal(await dialog.getByRole("button", { name: "Acknowledge current state" }).count(), 0);
    assert.match(await section.textContent() ?? "", /1 workbook changes require review/);
    await dialog.getByRole("button", { name: "Inspect", exact: true }).click();
    await dialog.getByRole("alert").waitFor();
    assert.match(await section.textContent() ?? "", /1 workbook changes require review/);
  } finally { await opened.finish(); }
});

void test("unidentified WPS can select completed history; obsolete deletion requires confirmation", async () => {
  const opened = await openBrowserPage(server, { path: "src/ui-gallery.html", viewport: { width: 350, height: 780 } });
  try {
    const { page } = opened;
    await page.locator('[data-gallery="durable-storage"]').getByRole("button", { name: "Retained history and storage" }).click();
    const dialog = page.getByRole("dialog");
    const retained = dialog.locator("section").filter({ hasText: "Budget.xlsx" });
    assert.match(await dialog.textContent() ?? "", /Identity not durable/);
    assert.equal(await dialog.getByRole("button", { name: "Delete obsolete generation…" }).count(), 1);
    await retained.getByRole("button", { name: "Inspect sessions" }).click();
    await retained.getByText("Pending execution — cannot copy").waitFor();
    assert.equal(await retained.getByRole("button", { name: "Copy completed history into current scope" }).count(), 1);
    await retained.getByRole("button", { name: "Copy completed history into current scope" }).click();
    await page.locator('[data-gallery="durable-storage"]').getByText("Completed history copied into current scope. No tasks resumed.").waitFor({ state: "attached" });
    await retained.getByRole("button", { name: "Delete obsolete generation…" }).click();
    await retained.getByRole("button", { name: "Cancel", exact: true }).click();
    assert.equal(await retained.count(), 1);
    await retained.getByRole("button", { name: "Delete obsolete generation…" }).click();
    await retained.getByRole("button", { name: "Confirm permanent deletion" }).click();
    await retained.waitFor({ state: "detached" });
    await page.keyboard.press("Escape");
    await page.locator('[data-gallery="durable-storage"]').getByRole("button", { name: "Retained history and storage" }).click();
    assert.equal(await page.getByRole("dialog").getByText("Budget.xlsx", { exact: true }).count(), 0);
  } finally { await opened.finish(); }
});
