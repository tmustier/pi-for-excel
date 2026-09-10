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

async function withFilesPage(paths: readonly string[], run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, { viewport: { width: 520, height: 760 } });
  try {
    await opened.page.evaluate(`
      (async () => {
        const { getFilesWorkspace } = await import("/src/files/workspace.ts");
        const workspace = getFilesWorkspace();
        for (const path of ${JSON.stringify(paths)}) {
          await workspace.writeTextFile(path, "browser fixture", "text/plain");
        }
      })()
    `);
    await enterCommand(opened.page, "/files");
    await opened.page.getByRole("heading", { name: "Files" }).waitFor({ state: "visible" });
    await opened.page.locator(".pi-files-section-group").first().waitFor({ state: "visible" });
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

async function withConnectedFilesPage(paths: readonly string[], run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, {
    viewport: { width: 520, height: 760 },
    prepareContext: async (context) => {
      await context.addInitScript(() => {
        const localFileHandle = {
          kind: "file",
          name: "browser-local.txt",
          getFile: () => Promise.resolve(new File(["local fixture"], "browser-local.txt", {
            type: "text/plain",
            lastModified: 1,
          })),
        };
        const root = {
          kind: "directory",
          name: "Project Docs",
          queryPermission: () => Promise.resolve("granted"),
          requestPermission: () => Promise.resolve("granted"),
          entries() {
            let emitted = false;
            return {
              [Symbol.asyncIterator]() {
                return this;
              },
              next() {
                if (emitted) return Promise.resolve({ done: true });
                emitted = true;
                return Promise.resolve({ done: false, value: ["browser-local.txt", localFileHandle] });
              },
            };
          },
        };
        Object.defineProperty(window, "showDirectoryPicker", {
          configurable: true,
          value: () => Promise.resolve(root),
        });
      });
    },
  });
  try {
    await opened.page.evaluate(`
      (async () => {
        const { getFilesWorkspace } = await import("/src/files/workspace.ts");
        const workspace = getFilesWorkspace();
        for (const path of ${JSON.stringify(paths)}) {
          await workspace.writeTextFile(path, "browser fixture", "text/plain");
        }
        await workspace.connectNativeDirectory();
      })()
    `);
    await enterCommand(opened.page, "/files");
    await opened.page.getByRole("heading", { name: "Files" }).waitFor({ state: "visible" });
    await opened.page.locator(".pi-files-section-group").first().waitFor({ state: "visible" });
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

function visibleFixtureNames(page: Page): Promise<string[]> {
  return page.locator(".pi-files-item__name").evaluateAll((elements) =>
    elements.map((element) => element.textContent ?? "").filter((name) => name.startsWith("browser-")));
}

const filteringRows = [
  {
    name: "trims and lowercases the entered filter",
    query: "  NOTES/BROWSER-QU  ",
    expected: ["browser-Quarterly-Plan.md"],
  },
  {
    name: "matches case-insensitive path substrings",
    query: "BROWSER-QUARTERLY",
    expected: ["browser-Quarterly-Plan.md"],
  },
] as const;

for (const row of filteringRows) {
  void test(`Files filtering ${row.name}`, async () => {
    await withFilesPage([
      "notes/browser-index.md",
      "notes/browser-Quarterly-Plan.md",
      "imports/browser-raw.csv",
    ], async (page) => {
      await page.getByPlaceholder("Filter files…").fill(row.query);
      assert.deepEqual(await visibleFixtureNames(page), row.expected);
    });
  });
}

void test("Files filtering leaves every file visible for whitespace", async () => {
  await withFilesPage([
    "browser-alpha.txt",
    "browser-beta.txt",
  ], async (page) => {
    await page.getByPlaceholder("Filter files…").fill("   ");
    assert.deepEqual(await visibleFixtureNames(page), ["browser-beta.txt", "browser-alpha.txt"]);
  });
});

void test("Files groups entries into visible category and folder sections", async () => {
  await withFilesPage([
    "browser-report.xlsx",
    "browser-data/raw.csv",
    "notes/browser-summary.md",
    "skills/browser-skill/SKILL.md",
  ], async (page) => {
    const text = (await page.locator(".pi-files-section-group").allInnerTexts()).join("\n");
    assert.match(text, /YOUR FILES[\s\S]*browser-data[\s\S]*PI'S NOTES[\s\S]*SKILLS[\s\S]*browser-skill/);
  });
});

void test("Files displays source badges with their visible priority", async () => {
  await withConnectedFilesPage([
    "notes/browser-today.md",
    "browser-upload.txt",
  ], async (page) => {
    const badgeFor = (name: string) => page.locator(".pi-files-item", { hasText: name }).locator(".pi-overlay-badge");
    assert.deepEqual(
      [
        await badgeFor("browser-today.md").textContent(),
        await badgeFor("browser-local.txt").textContent(),
        await page.locator(".pi-files-item--muted .pi-overlay-badge").first().textContent(),
        await badgeFor("browser-upload.txt").count(),
      ],
      ["Agent", "Folder", "Read only", 0],
    );
  });
});

void test("Files details identify each visible file source", async () => {
  await withConnectedFilesPage([
    "notes/browser-source-note.md",
    "browser-source-upload.txt",
  ], async (page) => {
    const sources: string[] = [];
    for (const item of [
      page.locator(".pi-files-item", { hasText: "browser-source-upload.txt" }),
      page.locator(".pi-files-item", { hasText: "browser-source-note.md" }),
      page.locator(".pi-files-item", { hasText: "browser-local.txt" }),
      page.locator(".pi-files-item--muted").first(),
    ]) {
      await item.click();
      sources.push(await page.locator(".pi-overlay-subtitle").last().innerText());
      await page.getByRole("button", { name: "Back to file list" }).click();
    }
    assert.deepEqual(sources.map((source) => source.split(" · ")[2]), [
      "Uploaded",
      "Written by agent",
      "Local file",
      "Pi documentation",
    ]);
  });
});

void test("Files exposes an enabled connect-folder action when the host supports it", async () => {
  const opened = await openTaskpane(env, {
    viewport: { width: 520, height: 760 },
    prepareContext: async (context) => {
      await context.addInitScript(() => {
        Object.defineProperty(window, "showDirectoryPicker", {
          configurable: true,
          value: () => Promise.reject(new DOMException("cancelled", "AbortError")),
        });
      });
    },
  });
  try {
    await enterCommand(opened.page, "/files");
    const button = opened.page.getByRole("button", { name: "Connect folder" });
    await button.waitFor({ state: "visible" });
    assert.equal(await button.isEnabled(), true);
  } finally {
    await opened.finish();
  }
});
