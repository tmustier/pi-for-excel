import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { enterCommand, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

const checkpoints = [
  { id: "cp1", at: 1_000, toolName: "write_cells", address: "Sheet1!A1:B10", changedCount: 5 },
  { id: "cp2", at: 2_000, toolName: "format_cells", address: "Sheet1!C1:C20", changedCount: 4 },
  { id: "cp3", at: 3_000, toolName: "write_cells", address: "Sheet2!A1:A5", changedCount: 3 },
  { id: "cp4", at: 4_000, toolName: "restore_snapshot", address: "Sheet1!A1:B10", changedCount: 2, restoredFromSnapshotId: "cp1" },
  { id: "cp5", at: 5_000, toolName: "modify_structure", address: "Sheet3", changedCount: 1 },
] as const;

async function withRecoveryPage(
  run: (page: Page) => Promise<void>,
  entries: readonly object[] = checkpoints,
): Promise<void> {
  const opened = await openTaskpane(env, { viewport: { width: 520, height: 760 } });
  try {
    await opened.page.evaluate(`
      (async () => {
        const { configureSettingsPages } = await import("/src/commands/builtins/settings-pages/index.ts");
        const entries = ${JSON.stringify(entries)};
        configureSettingsPages({
          backups: {
            loadCheckpoints: () => Promise.resolve(entries),
            onRestore: () => Promise.resolve(),
            onDelete: () => Promise.resolve(true),
            onClear: () => Promise.resolve(entries.length),
          },
        });
      })()
    `);
    await enterCommand(opened.page, "/history");
    await opened.page.getByRole("heading", { name: "Backups" }).waitFor({ state: "visible" });
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

function visibleIds(page: Page): Promise<string[]> {
  return page.locator(".pi-recovery-item__meta").evaluateAll((elements) =>
    elements.map((element) => element.textContent?.split("#")[1] ?? ""));
}

async function search(page: Page, query: string): Promise<void> {
  await page.getByPlaceholder("Search backups…").fill(query);
  await page.waitForTimeout(250);
}

void test("Recovery shows every backup newest first by default", async () => {
  await withRecoveryPage(async (page) => {
    assert.deepEqual(await visibleIds(page), ["cp5", "cp4", "cp3", "cp2", "cp1"]);
  });
});

void test("Recovery can sort backups oldest first", async () => {
  await withRecoveryPage(async (page) => {
    await page.getByRole("button", { name: "↓ Newest" }).click();
    assert.deepEqual(await visibleIds(page), ["cp1", "cp2", "cp3", "cp4", "cp5"]);
  });
});

void test("Recovery tool selection shows only that action", async () => {
  await withRecoveryPage(async (page) => {
    await page.locator(".pi-recovery-filter-select").selectOption("write_cells");
    assert.deepEqual(await visibleIds(page), ["cp3", "cp1"]);
  });
});

const searchRows = [
  { name: "backup id", query: "cp3", expected: ["cp3"] },
  { name: "address without case sensitivity", query: "sHeEt2", expected: ["cp3"] },
  { name: "visible tool label", query: "Modify structure", expected: ["cp5"] },
] as const;

for (const row of searchRows) {
  void test(`Recovery search matches ${row.name}`, async () => {
    await withRecoveryPage(async (page) => {
      await search(page, row.query);
      assert.deepEqual(await visibleIds(page), row.expected);
    });
  });
}

void test("Recovery combines search and tool selection", async () => {
  await withRecoveryPage(async (page) => {
    await search(page, "Sheet1");
    await page.locator(".pi-recovery-filter-select").selectOption("write_cells");
    assert.deepEqual(await visibleIds(page), ["cp1"]);
  });
});

void test("Recovery treats whitespace search as empty", async () => {
  await withRecoveryPage(async (page) => {
    await search(page, "   ");
    assert.deepEqual(await visibleIds(page), ["cp5", "cp4", "cp3", "cp2", "cp1"]);
  });
});

void test("Recovery renders an empty result when no backup matches", async () => {
  await withRecoveryPage(async (page) => {
    await search(page, "nonexistent");
    assert.equal(await page.locator(".pi-recovery-list").innerText(), "No backups match the current filter.");
  });
});

void test("Recovery action choices show counts from the loaded backups", async () => {
  await withRecoveryPage(async (page) => {
    assert.deepEqual(await page.locator(".pi-recovery-filter-select option").allTextContents(), [
      "All actions (5)",
      "Write (2)",
      "Format cells (1)",
      "Modify structure (1)",
      "Restore (1)",
    ]);
  });
});

void test("Recovery action choices omit actions with no backups", async () => {
  await withRecoveryPage(async (page) => {
    assert.equal(await page.locator('.pi-recovery-filter-select option[value="python_transform_range"]').count(), 0);
  });
});

void test("Recovery with no backups exposes no action filter", async () => {
  await withRecoveryPage(async (page) => {
    assert.equal(await page.locator(".pi-recovery-search-row").isHidden(), true);
  }, []);
});
