import assert from "node:assert/strict";
import { test } from "node:test";

import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { deferred } from "./fixtures/workbook-deferred.ts";

void test("throwing listeners cannot strand writes, suppress other listeners, or change results", async () => {
  const coordinator = createWorkbookCoordinator();
  const events: string[] = [];
  coordinator.subscribe(() => { throw new Error("subscriber failed"); });
  coordinator.subscribe((event) => { events.push(`${event.context.opId}:${event.type}`); });
  const context = (opId: string) => ({ workbookId: "book", sessionId: "tab", opId });
  const release = deferred<void>();
  const first = coordinator.runWrite(context("first"), async () => { await release.promise; return 7; });
  const failed = coordinator.runWrite(context("failed"), () => Promise.reject(new Error("host failed")));
  const rejected = assert.rejects(failed, /host failed/u);
  const last = coordinator.runWrite(context("last"), () => Promise.resolve(9));
  release.resolve();
  const outcomes = Promise.all([first, rejected, last]);
  await assert.doesNotReject(outcomes, "listener exceptions must not reject workbook operations");
  const [firstResult, , lastResult] = await outcomes;
  assert.deepEqual(firstResult, { result: 7, revision: 1 });
  assert.deepEqual(lastResult, { result: 9, revision: 2 });
  assert.deepEqual(events, [
    "first:started", "failed:queued", "last:queued", "first:completed",
    "failed:started", "failed:failed", "last:started", "last:completed",
  ]);
  await coordinator.runRead(context("read"), () => Promise.resolve(11));
  await assert.rejects(coordinator.runRead(context("read-error"), () => Promise.reject(new Error("read failed"))), /read failed/u);
  assert.equal(coordinator.getSnapshot("book").activeWrite, null);
  assert.equal(coordinator.getSnapshot("book").queuedWrites, 0);
});
