import assert from "node:assert/strict";
import { test } from "node:test";

import { createExtensionAPI, loadExtension } from "../src/commands/extension-api.ts";

import {
  collectActivationCleanups,
  createLoadedExtensionHandle,
  getExtensionActivator,
  getExtensionDeactivator,
} from "../src/commands/extension-loader.ts";

void test("collectActivationCleanups accepts undefined, function, and function arrays", () => {
  const cleanupA = () => undefined;
  const cleanupB = () => undefined;

  assert.deepEqual(collectActivationCleanups(undefined), []);
  assert.deepEqual(collectActivationCleanups(cleanupA), [cleanupA]);
  assert.deepEqual(collectActivationCleanups([cleanupA, cleanupB]), [cleanupA, cleanupB]);
});

void test("collectActivationCleanups rejects invalid return values", () => {
  assert.throws(
    () => collectActivationCleanups("not-valid"),
    /activate\(api\) must return void, a cleanup function, or an array of cleanup functions/,
  );

  assert.throws(
    () => collectActivationCleanups([() => undefined, "bad"]),
    /activate\(api\) returned an invalid cleanup entry; expected a function/,
  );
});

void test("getExtensionActivator resolves activate then default", () => {
  const activate = () => undefined;
  const fallback = () => undefined;

  assert.equal(getExtensionActivator({ activate }), activate);
  assert.equal(getExtensionActivator({ default: fallback }), fallback);
  assert.equal(getExtensionActivator({}), null);
});

void test("getExtensionDeactivator resolves optional deactivate", () => {
  const deactivate = () => undefined;

  assert.equal(getExtensionDeactivator({ deactivate }), deactivate);
  assert.equal(getExtensionDeactivator({}), null);
  assert.equal(getExtensionDeactivator(null), null);
});

void test("createLoadedExtensionHandle executes cleanup in reverse order and only once", async () => {
  const calls: string[] = [];

  const handle = createLoadedExtensionHandle([
    () => {
      calls.push("cleanup-1");
    },
    () => {
      calls.push("cleanup-2");
    },
  ], () => {
    calls.push("module-deactivate");
  });

  await handle.deactivate();
  await handle.deactivate();

  assert.deepEqual(calls, ["cleanup-2", "cleanup-1", "module-deactivate"]);
});

void test("createLoadedExtensionHandle aggregates cleanup failures", async () => {
  const handle = createLoadedExtensionHandle([
    () => {
      throw new Error("cleanup-a");
    },
  ], () => {
    throw new Error("cleanup-b");
  });

  await assert.rejects(
    handle.deactivate(),
    /Extension cleanup failed:\n- cleanup-a\n- cleanup-b/,
  );
});

void test("concurrent deactivation callers wait for the same teardown", async () => {
  let release = (): void => {};
  const blocked = new Promise<void>((resolve) => { release = resolve; });
  let calls = 0;
  const api = createExtensionAPI({ getAgent: () => { throw new Error("unused"); } });
  const handle = await loadExtension(api, () => [async () => { calls += 1; await blocked; }]);
  const first = handle.deactivate();
  const second = handle.deactivate();
  assert.equal(first, second, "concurrent callers must share the in-flight teardown promise");
  release();
  await Promise.all([first, second]);
  assert.equal(calls, 1);
});

void test("deactivation retries only failed cleanups and failed module teardown", async () => {
  const calls: string[] = [];
  let fail = true;
  const handle = createLoadedExtensionHandle([
    () => { calls.push("success"); },
    () => { calls.push("retry"); if (fail) throw new Error("retry cleanup"); },
  ], () => { calls.push("module"); if (fail) throw new Error("retry module"); });
  await assert.rejects(handle.deactivate(), /retry cleanup[\s\S]*retry module/);
  fail = false;
  await handle.deactivate();
  await handle.deactivate();
  assert.deepEqual(calls, ["retry", "success", "module", "retry", "module"]);
});

void test("loaded extensions retry failed cleanup without replaying successful cleanup", async () => {
  const api = createExtensionAPI({ getAgent: () => { throw new Error("unused"); } });
  const calls: string[] = [];
  let failed = true;
  const handle = await loadExtension(api, () => [
    () => { calls.push("success"); },
    () => { calls.push("retry"); if (failed) throw new Error("cleanup unavailable"); },
  ]);
  await assert.rejects(handle.deactivate(), /cleanup unavailable/);
  failed = false;
  await handle.deactivate();
  await handle.deactivate();
  assert.deepEqual(calls, ["retry", "success", "retry"]);
});
