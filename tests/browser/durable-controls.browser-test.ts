import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { join } from "node:path";
import { openBrowserPage, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let server: TaskpaneServer;
before(async () => { server = await startTaskpaneServer(); });
after(async () => { await server.close(); });

const TASKPANE = { width: 350, height: 780 };

void test("review summarizes the record, keeps forensics collapsed, and the action is reachable in a 350x780 taskpane", async () => {
  const opened = await openBrowserPage(server, { path: "src/ui-gallery.html", viewport: TASKPANE });
  const { page } = opened;
  try {
    const section = page.locator('[data-gallery="durable-conflict"]');
    assert.match(await section.textContent() ?? "", /Workbook changes awaiting review: 1\./);
    assert.doesNotMatch(await section.textContent() ?? "", /1 workbook changes/);
    await section.getByRole("button", { name: "Inspect", exact: true }).click();
    const dialog = page.getByRole("dialog");
    const acknowledge = dialog.getByRole("button", { name: "Acknowledge current state" });
    await acknowledge.waitFor();
    const summary = await dialog.innerText();
    assert.match(summary, /Operation\s+write_cells/);
    assert.match(summary, /Target\s+Budget!C4:D4/);
    assert.match(summary, /Interrupted after the workbook call started/);
    assert.match(summary, /Outcome uncertain/);
    assert.match(summary, /Original\s+C4:D4: 1000, =C4\*1\.2/);
    assert.match(summary, /Intended\s+C4:D4: 1200, =C4\*1\.2/);
    assert.match(summary, /Now\s+C4:D4: 1175, =C4\*1\.2/);
    assert.match(summary, /never writes to the workbook, and does not prove the original operation succeeded/);
    // Forensic data is present but collapsed: no token, GUID or record JSON in the default view.
    assert.doesNotMatch(summary, /observationToken|gallery-observation|4F1C2B9E|"version"/);
    const evidence = dialog.getByRole("group").filter({ hasText: "Technical evidence" });
    assert.equal(await evidence.evaluate((node) => node instanceof HTMLDetailsElement && !node.open), true);
    // The decision fits the taskpane without scrolling past evidence.
    const box = await acknowledge.boundingBox();
    assert.ok(box && box.y + box.height <= TASKPANE.height, `acknowledge bottom ${box ? box.y + box.height : "missing"} exceeds ${TASKPANE.height}`);
    await acknowledge.focus();
    await page.keyboard.press("Tab");
    await page.keyboard.press("Tab");
    assert.equal(await page.evaluate(() => document.activeElement?.textContent), "Technical evidence");
    await page.keyboard.press("Enter");
    await dialog.getByText(/"observationToken": "gallery-observation-1"/).waitFor();
    await dialog.getByText(/mutation:7c1e0f4a-2b9d-4e63-a8f5-91d3c6b20e47/).first().waitFor();
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  } finally { await opened.finish(); }
});

void test("stale acknowledgement is refused, reinspected, and needs a second explicit acknowledgement", async () => {
  const opened = await openBrowserPage(server, { path: "src/ui-gallery.html", viewport: TASKPANE });
  const { page } = opened;
  try {
    const section = page.locator('[data-gallery="durable-conflict"]');
    await section.getByRole("button", { name: "Inspect", exact: true }).click();
    const dialog = page.getByRole("dialog");
    await dialog.getByRole("button", { name: "Acknowledge current state" }).click();
    await dialog.getByRole("alert").waitFor();
    await dialog.getByText("C4:D4: 1190, =C4*1.2").waitFor();
    assert.match(await dialog.innerText(), /Acceptance invalidated/);
    assert.match(await section.textContent() ?? "", /Workbook changes awaiting review: 1\./);
    // Focus lands on the new result, never on the new acknowledgement.
    assert.equal(await page.evaluate(() => document.activeElement?.textContent), "Outcome uncertain");
    const again = dialog.getByRole("button", { name: "Acknowledge current state" });
    const box = await again.boundingBox();
    assert.ok(box && box.y + box.height <= TASKPANE.height);
    await again.click();
    await dialog.waitFor({ state: "detached" });
    await section.getByText("No workbook changes awaiting review.", { exact: true }).waitFor({ state: "attached" });
    assert.match(await section.textContent() ?? "", /original success is not asserted/);
    assert.equal(await section.getByRole("button", { name: "Refresh reviews" }).evaluate((node) => node === document.activeElement), true);
    const pendingInspect = page.locator('[data-gallery="durable-pending"]').getByRole("button", { name: "Inspect", exact: true });
    await pendingInspect.click();
    await dialog.getByText("Intended result found").waitFor();
    await page.keyboard.press("Escape");
    await dialog.waitFor({ state: "detached" });
    assert.equal(await pendingInspect.evaluate((node) => node === document.activeElement), true);
  } finally { await opened.finish(); }
});

for (const height of [780, 680]) {
  for (const scenario of ["pending", "conflict"]) {
    void test(`review decisions stay visible with long reasons, stale refusal and expanded evidence: ${scenario} at 350x${height}`, async () => {
      const opened = await openBrowserPage(server, { path: "src/ui-gallery.html", viewport: { width: 350, height } });
      const { page } = opened;
      try {
        const section = page.locator(`[data-gallery="durable-${scenario}"]`);
        await section.getByRole("button", { name: "Inspect", exact: true }).click();
        const dialog = page.getByRole("dialog");
        const acknowledge = dialog.getByRole("button", { name: "Acknowledge current state" });
        await acknowledge.waitFor();
        // Both real gallery outcomes include the full, multi-line evidence limits;
        // the applied case also has a separate reason, as in the native review.
        assert.match(await dialog.innerText(), /persistent row\/column identity, and never authorizes replay/);
        const assertDecisionsVisible = async (state: string): Promise<void> => {
          const screenshotDir = process.env.PI_DURABLE_SCREENSHOT_DIR;
          if (screenshotDir) await page.screenshot({ path: join(screenshotDir, `${scenario}-${height}-${state}.png`) });
          for (const name of ["Acknowledge current state", "Check again"]) {
            const button = dialog.getByRole("button", { name, exact: true });
            const visible = await button.evaluate((node) => {
              const box = node.getBoundingClientRect();
              const card = node.closest(".pi-durable-dialog")?.getBoundingClientRect();
              return card && box.top >= Math.max(0, card.top) && box.bottom <= Math.min(innerHeight, card.bottom)
                && box.left >= 0 && box.right <= innerWidth
                && node.contains(document.elementFromPoint(box.x + box.width / 2, box.y + box.height / 2));
            });
            assert.equal(visible, true, `${name} must be visible and unobscured (${scenario}, ${height}, ${state})`);
          }
          assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
        };
        const checkEvidence = async (state: string): Promise<void> => {
          const warning = dialog.getByText(/^Acknowledging accepts the workbook as it is now/);
          await warning.evaluate((node) => node.scrollIntoView({ block: "center" }));
          const warningReadable = await warning.evaluate((node) => {
            const box = node.getBoundingClientRect();
            return node.contains(document.elementFromPoint(box.x + box.width / 2, box.top + 4))
              && node.contains(document.elementFromPoint(box.x + box.width / 2, box.bottom - 4));
          });
          assert.equal(warningReadable, true, "the full acknowledgement warning remains reachable above the actions");
          await assertDecisionsVisible(`${state}-warning`);
          const evidence = dialog.getByRole("group").filter({ hasText: "Technical evidence" });
          assert.equal(await evidence.evaluate((node) => node instanceof HTMLDetailsElement && !node.open), true);
          const summary = evidence.locator("summary");
          await summary.click();
          assert.equal(await evidence.evaluate((node) => node instanceof HTMLDetailsElement && node.open), true);
          await assertDecisionsVisible(`${state}-expanded`);
          const observation = evidence.locator("pre").last();
          assert.match(await observation.textContent() ?? "", /"observationToken": "gallery-/);
          // Read right to the end of the full observation, not just the disclosure.
          await observation.evaluate((node) => node.scrollIntoView({ block: "end" }));
          const endVisible = await observation.evaluate((node) => {
            const box = node.getBoundingClientRect();
            return box.bottom > 0 && box.bottom <= innerHeight
              && node.contains(document.elementFromPoint(box.x + box.width / 2, box.bottom - 4));
          });
          assert.equal(endVisible, true, "expanded evidence must remain readable through its final line");
          await assertDecisionsVisible(`${state}-evidence-end`);
        };
        await assertDecisionsVisible("fresh");
        if (scenario === "conflict") {
          await acknowledge.click();
          await dialog.getByRole("alert").waitFor();
          await dialog.getByText("C4:D4: 1190, =C4*1.2").waitFor();
          assert.match(await dialog.innerText(), /Acceptance invalidated/);
          assert.equal(await page.evaluate(() => document.activeElement?.textContent), "Outcome uncertain");
          assert.match(await section.textContent() ?? "", /Workbook changes awaiting review: 1\./);
          await assertDecisionsVisible("stale");
          await checkEvidence("stale");
        } else {
          await checkEvidence("fresh");
        }
        await dialog.getByRole("button", { name: "Check again", exact: true }).click();
        await acknowledge.waitFor();
        await assertDecisionsVisible("rechecked");
        await acknowledge.click();
        await dialog.waitFor({ state: "detached" });
        await section.getByText("No workbook changes awaiting review.", { exact: true }).waitFor({ state: "attached" });
      } finally { await opened.finish(); }
    });
  }
}

void test("inspection errors retain pending and offer no acknowledgement", async () => {
  const opened = await openBrowserPage(server, { path: "src/ui-gallery.html", viewport: TASKPANE });
  try {
    const section = opened.page.locator('[data-gallery="durable-error"]');
    await section.getByRole("button", { name: "Inspect", exact: true }).click();
    const dialog = opened.page.getByRole("dialog");
    await dialog.getByRole("alert").waitFor();
    await dialog.getByText(/could not be checked/).waitFor();
    assert.match(await dialog.innerText(), /Operation\s+Change recorded by an earlier version/);
    assert.match(await dialog.innerText(), /Target\s+Not recorded/);
    assert.equal(await dialog.getByText("Technical evidence").count(), 1);
    assert.equal(await dialog.getByRole("button", { name: "Acknowledge current state" }).count(), 0);
    assert.match(await section.textContent() ?? "", /Workbook changes awaiting review: 1\./);
    await dialog.getByRole("button", { name: "Check again", exact: true }).click();
    await dialog.getByRole("alert").waitFor();
    assert.equal(await dialog.getByRole("button", { name: "Acknowledge current state" }).count(), 0);
    assert.match(await section.textContent() ?? "", /Workbook changes awaiting review: 1\./);
  } finally { await opened.finish(); }
});

void test("unidentified WPS can select completed history; obsolete deletion requires confirmation", async () => {
  const opened = await openBrowserPage(server, { path: "src/ui-gallery.html", viewport: TASKPANE });
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
