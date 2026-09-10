/* eslint-disable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access -- Playwright evaluates Vite-loaded application modules whose browser runtime types are unavailable to the Node test checker. */
import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { enterCommand, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, { viewport: { width: 420, height: 720 }, clientId: "ui-dialogs-browser" });
  try { await run(opened.page); } finally { await opened.finish(); }
}

async function importFile(page: Page, name: string, contents = "hello", type = "text/plain"): Promise<void> {
  await page.evaluate(async ({ name, contents, type }) => {
    const { getFilesWorkspace } = await import("/src/files/workspace.ts");
    await getFilesWorkspace().importFiles([new File([contents], name, { type, lastModified: 1 })]);
  }, { name, contents, type });
}

async function openFileDetail(page: Page, name: string): Promise<void> {
  await enterCommand(page, "/files");
  await page.getByText(name, { exact: true }).click();
  await page.locator(".pi-files-detail-actions").waitFor({ state: "visible" });
}

void test("the Files rename dialog confirms a new path", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "plan.md");
    await openFileDetail(page, "plan.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    await dialog.getByRole("textbox").fill("archive/final-plan");
    await dialog.getByRole("button", { name: "Rename" }).click();
    await page.getByRole("heading", { name: "final-plan.md" }).waitFor({ state: "visible" });
    assert.equal(await page.getByRole("heading", { name: "final-plan.md" }).innerText(), "final-plan.md");
  });
});

void test("the Files rename dialog cancels without changing the file", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "cancel-me.md");
    await openFileDetail(page, "cancel-me.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    await dialog.getByRole("textbox").fill("changed.md");
    await dialog.getByRole("button", { name: "Cancel", exact: true }).last().click();
    await dialog.waitFor({ state: "detached" });
    assert.equal(await page.getByRole("heading", { name: "cancel-me.md" }).innerText(), "cancel-me.md");
  });
});

void test("Escape leaves text entry then dismisses the dialog and permits a fresh mount", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "escape.md");
    await openFileDetail(page, "escape.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    await page.waitForFunction(() => document.activeElement?.matches("#pi-text-input-dialog-overlay input") === true);
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await dialog.waitFor({ state: "detached" });
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    assert.equal(await page.locator("#pi-text-input-dialog-overlay").count(), 1);
  });
});

void test("an overlay exposes modal dialog semantics and focuses its text input", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "focus.md");
    await openFileDetail(page, "focus.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    assert.deepEqual({
      role: await dialog.getAttribute("role"),
      modal: await dialog.getAttribute("aria-modal"),
      focused: await dialog.getByRole("textbox").evaluate((input) => input === document.activeElement),
    }, { role: "dialog", modal: "true", focused: true });
  });
});

void test("read-only built-in files expose only copy and download actions", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/files");
    const builtIn = page.locator(".pi-files-item--muted").first();
    await builtIn.click();
    const actions = page.locator(".pi-files-detail-actions button");
    await actions.last().waitFor({ state: "visible" });
    assert.deepEqual(await actions.allTextContents(), ["Copy content", "Download"]);
  });
});

for (const row of [
  { name: "unsafe.html", type: "text/html", expected: "application/octet-stream" },
  { name: "safe.txt", type: "text/plain", expected: "text/plain" },
] as const) {
  void test(`Files Open uses a safe blob type for ${row.type}`, async () => {
    await withTaskpane(async (page) => {
      await importFile(page, row.name, "sample", row.type);
      await page.evaluate(() => {
        Object.defineProperty(window, "open", { configurable: true, value: () => null });
        const original = URL.createObjectURL.bind(URL);
        Object.defineProperty(URL, "createObjectURL", { configurable: true, value: (blob: Blob) => {
          document.body.dataset.openedBlobType = blob.type;
          return original(blob);
        } });
      });
      await openFileDetail(page, row.name);
      await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Open ↗" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
      await page.waitForFunction(() => document.body.dataset.openedBlobType !== undefined);
      assert.equal(await page.locator("body").getAttribute("data-opened-blob-type"), row.expected);
    });
  });
}

void test("the Files footer reports count, total size, and active backend", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "sized.txt", "x".repeat(1024));
    await enterCommand(page, "/files");
    const footer = page.locator(".pi-files-footer");
    await page.waitForFunction(() => (document.querySelector(".pi-files-footer")?.textContent?.length ?? 0) > 0);
    assert.match(await footer.innerText(), /\d+ files? · [\d.]+ (?:KB|B) · Sandboxed workspace/);
  });
});
