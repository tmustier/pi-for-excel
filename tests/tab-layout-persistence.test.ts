import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createTabLayoutPersistence,
} from "../src/taskpane/tab-layout-persistence.ts";
import {
  loadWorkbookTabLayout,
  saveWorkbookTabLayout,
  type WorkbookTabLayout,
  workbookTabLayoutKey,
} from "../src/taskpane/tab-layout.ts";

const SAMPLE_LAYOUT: WorkbookTabLayout = {
  sessionIds: ["session-a", "session-b"],
  activeSessionId: "session-b",
};

class StartupLayoutSettings {
  private readonly values = new Map<string, unknown>();
  private failRead = false;

  get(key: string): Promise<unknown> {
    if (this.failRead) {
      this.failRead = false;
      return Promise.reject(new Error("seeded layout read failure"));
    }
    return Promise.resolve(this.values.get(key) ?? null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  delete(key: string): Promise<void> {
    this.values.delete(key);
    return Promise.resolve();
  }

  seed(value: unknown, failRead = false): void {
    this.values.set(workbookTabLayoutKey("wb-1"), value);
    this.failRead = failRead;
  }

  stored(): unknown {
    return this.values.get(workbookTabLayoutKey("wb-1")) ?? null;
  }
}

async function persistStartupFallback(settings: StartupLayoutSettings): Promise<void> {
  let readSucceeded = false;
  try {
    await loadWorkbookTabLayout(settings, "wb-1");
    readSucceeded = true;
  } catch {
    // Startup falls back to a fresh runtime, but must preserve unread storage.
  }

  const fallback: WorkbookTabLayout = {
    sessionIds: ["fresh-session"],
    activeSessionId: "fresh-session",
  };
  const controller = createTabLayoutPersistence({
    resolveWorkbookId: () => Promise.resolve("wb-1"),
    saveLayout: (workbookId, layout) => saveWorkbookTabLayout(settings, workbookId, layout),
  });
  controller.enable(readSucceeded ? undefined : fallback);
  controller.persist(fallback);
  await controller.flush();
}

void test("startup preserves an unread layout but replaces a corrupt layout", async () => {
  const persistedLayout = {
    sessionIds: ["persisted-session"],
    activeSessionId: "persisted-session",
  };
  const unreadable = new StartupLayoutSettings();
  unreadable.seed(persistedLayout, true);
  await persistStartupFallback(unreadable);
  assert.deepEqual(unreadable.stored(), persistedLayout);

  const corrupt = new StartupLayoutSettings();
  corrupt.seed({ sessionIds: "corrupt" });
  await persistStartupFallback(corrupt);
  assert.deepEqual(corrupt.stored(), {
    sessionIds: ["fresh-session"],
    activeSessionId: "fresh-session",
  });
});

void test("tab layout persistence is disabled until enabled", async () => {
  const saves: Array<{ workbookId: string | null; layout: WorkbookTabLayout }> = [];

  const controller = createTabLayoutPersistence({
    resolveWorkbookId: () => Promise.resolve("wb-1"),
    saveLayout: (workbookId, layout) => {
      saves.push({ workbookId, layout });
      return Promise.resolve();
    },
  });

  controller.persist(SAMPLE_LAYOUT);
  await controller.flush();

  assert.equal(saves.length, 0);
});

void test("tab layout persistence deduplicates same workbook+layout signature", async () => {
  const saves: Array<{ workbookId: string | null; layout: WorkbookTabLayout }> = [];

  const controller = createTabLayoutPersistence({
    resolveWorkbookId: () => Promise.resolve("wb-1"),
    saveLayout: (workbookId, layout) => {
      saves.push({ workbookId, layout });
      return Promise.resolve();
    },
  });

  controller.enable();
  controller.persist(SAMPLE_LAYOUT);
  controller.persist(SAMPLE_LAYOUT);
  await controller.flush();

  assert.equal(saves.length, 1);
  assert.equal(saves[0]?.workbookId, "wb-1");
});

void test("tab layout persistence does not dedupe across workbook ids", async () => {
  const saves: Array<{ workbookId: string | null; layout: WorkbookTabLayout }> = [];
  const workbookIds = ["wb-1", "wb-2"];
  let nextWorkbookIndex = 0;

  const controller = createTabLayoutPersistence({
    resolveWorkbookId: () => {
      const workbookId = workbookIds[nextWorkbookIndex] ?? null;
      nextWorkbookIndex += 1;
      return Promise.resolve(workbookId);
    },
    saveLayout: (workbookId, layout) => {
      saves.push({ workbookId, layout });
      return Promise.resolve();
    },
  });

  controller.enable();
  controller.persist(SAMPLE_LAYOUT);
  controller.persist(SAMPLE_LAYOUT);
  await controller.flush();

  assert.equal(saves.length, 2);
  assert.deepEqual(
    saves.map((save) => save.workbookId),
    ["wb-1", "wb-2"],
  );
});

void test("tab layout persistence keeps queue alive after save failure", async () => {
  const savedWorkbookIds: Array<string | null> = [];
  const warnings: string[] = [];
  let saveAttempts = 0;

  const controller = createTabLayoutPersistence({
    resolveWorkbookId: () => Promise.resolve("wb-1"),
    saveLayout: (workbookId, _layout) => {
      saveAttempts += 1;
      if (saveAttempts === 1) {
        return Promise.reject(new Error("disk full"));
      }
      savedWorkbookIds.push(workbookId);
      return Promise.resolve();
    },
    warn: (message, error) => {
      const suffix = error instanceof Error ? ` ${error.message}` : "";
      warnings.push(`${message}${suffix}`);
    },
  });

  controller.enable();
  controller.persist({
    sessionIds: ["session-a"],
    activeSessionId: "session-a",
  });
  controller.persist({
    sessionIds: ["session-a", "session-c"],
    activeSessionId: "session-c",
  });
  await controller.flush();

  assert.equal(saveAttempts, 2);
  assert.deepEqual(savedWorkbookIds, ["wb-1"]);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0] ?? "", /Failed to persist tab layout/);
});
