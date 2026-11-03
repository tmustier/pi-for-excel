import assert from "node:assert/strict";
import { test } from "node:test";

import { WorkbookChangeAuditLog } from "../src/audit/workbook-change-audit.ts";
import type { SettingsUpdater } from "../src/storage/local/settings-store.ts";
import { createInMemorySettingsStore } from "./fixtures/recovery-log.ts";
import { deferred } from "./fixtures/workbook-deferred.ts";

void test("clear invalidates an audit batch waiting to enter persistence", async () => {
  const backing = createInMemorySettingsStore();
  const flushStarted = deferred<void>();
  const releaseFlush = deferred<void>();
  let updates = 0;
  const settings: SettingsUpdater = {
    get: backing.get,
    set: backing.set,
    update: async (key, parse, change) => {
      updates += 1;
      if (updates === 1) {
        flushStarted.resolve();
        await releaseFlush.promise;
      }
      await backing.update(key, parse, change);
    },
  };
  const log = new WorkbookChangeAuditLog({
    settings,
    getWorkbookContext: () => Promise.resolve({ workbookId: "book", workbookName: "Book", source: "document.url" }),
    createId: () => "entry",
  });
  const append = log.append({
    toolName: "write_cells", toolCallId: "call", blocked: false,
    changedCount: 1, changes: [], executionMode: "safe",
  });
  await flushStarted.promise;
  await log.clear();
  releaseFlush.resolve();
  await append;
  assert.deepEqual(await log.list(), [], "clear must not resurrect the captured audit batch");
  const reloaded = new WorkbookChangeAuditLog({ settings });
  assert.deepEqual(await reloaded.list(), [], "cleared entries must remain absent after reload");
});
