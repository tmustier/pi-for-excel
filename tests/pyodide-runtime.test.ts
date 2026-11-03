import assert from "node:assert/strict";
import { test, type TestContext } from "node:test";
import {
  callPyodideRuntime,
  isPyodideLoaded,
  terminatePyodide,
} from "../src/python/pyodide-runtime.ts";
import { ToolStoppedError } from "../src/tools/tool-stopped.ts";
import type { PyodideWorkerRequest, PyodideWorkerMessage } from "../src/python/pyodide-worker.ts";

class WorkerHost {
  static instances: WorkerHost[] = [];
  requests: PyodideWorkerRequest[] = [];
  terminated = false;
  message: (event: { data: PyodideWorkerMessage }) => void = () => {};
  error: (event: { message: string }) => void = () => {};

  constructor() { WorkerHost.instances.push(this); }
  addEventListener(type: string, listener: typeof this.message & typeof this.error): void {
    if (type === "message") this.message = listener;
    if (type === "error") this.error = listener;
  }
  postMessage(request: PyodideWorkerRequest): void { this.requests.push(request); }
  terminate(): void { this.terminated = true; }
  complete(): void {
    const request = this.requests.at(-1);
    assert.ok(request);
    this.message({ data: { type: "loaded" } });
    this.message({ data: { id: request.id, ok: true, stdout: request.code, loadTimeMs: 1 } });
  }
}

function installHost(t: TestContext): void {
  const original = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  terminatePyodide();
  WorkerHost.instances = [];
  Reflect.set(globalThis, "Worker", WorkerHost);
  t.after(() => {
    terminatePyodide();
    if (original) Object.defineProperty(globalThis, "Worker", original);
    else Reflect.deleteProperty(globalThis, "Worker");
  });
}

void test("Pyodide dispatches one execution at a time and reports loaded only after initialization", async (t) => {
  installHost(t);
  const first = callPyodideRuntime({ code: "first" });
  const second = callPyodideRuntime({ code: "second" });
  const host = WorkerHost.instances[0];
  assert.ok(host);
  assert.deepEqual(host.requests.map((request) => request.code), ["first"]);
  assert.equal(isPyodideLoaded(), false);
  host.complete();
  assert.equal((await first).stdout, "first");
  assert.equal(isPyodideLoaded(), true);
  assert.deepEqual(host.requests.map((request) => request.code), ["first", "second"]);
  host.complete();
  assert.equal((await second).stdout, "second");
});

void test("Pyodide timeout settles the owning queue and cannot terminate a replacement worker", async (t) => {
  installHost(t);
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const first = callPyodideRuntime({ code: "first", timeout_ms: 10 });
  const second = callPyodideRuntime({ code: "second", timeout_ms: 20 });
  let secondSettled = false;
  void second.then(() => { secondSettled = true; });
  t.mock.timers.tick(10);
  assert.match((await first).error ?? "", /timed out/u);
  await Promise.resolve();
  assert.equal(secondSettled, true);
  assert.equal((await second).ok, false);
  const replacement = callPyodideRuntime({ code: "replacement", timeout_ms: 100 });
  const host = WorkerHost.instances[1];
  assert.ok(host);
  t.mock.timers.tick(20);
  assert.equal(host.terminated, false);
  host.complete();
  assert.equal((await replacement).ok, true);
});

void test("Pyodide abort terminates running Python and settles its queued requests", async (t) => {
  installHost(t);
  const controller = new AbortController();
  const first = callPyodideRuntime({ code: "running" }, controller.signal);
  const second = callPyodideRuntime({ code: "queued" });
  controller.abort();
  await assert.rejects(first, ToolStoppedError);
  assert.equal(WorkerHost.instances[0]?.terminated, true);
  assert.equal((await second).error, "Aborted");
});

void test("Pyodide abort of queued work does not interrupt running Python", async (t) => {
  installHost(t);
  const first = callPyodideRuntime({ code: "running" });
  const controller = new AbortController();
  const second = callPyodideRuntime({ code: "queued" }, controller.signal);
  controller.abort();
  await assert.rejects(second, ToolStoppedError);
  const host = WorkerHost.instances[0];
  assert.ok(host);
  assert.equal(host.requests.length, 1);
  assert.equal(host.terminated, false);
  host.complete();
  assert.equal((await first).ok, true);
});

void test("Pyodide ignores errors from a terminated worker after a replacement starts", async (t) => {
  installHost(t);
  const first = callPyodideRuntime({ code: "first" });
  const old = WorkerHost.instances[0];
  assert.ok(old);
  terminatePyodide();
  await first;
  const replacement = callPyodideRuntime({ code: "replacement" });
  const host = WorkerHost.instances[1];
  assert.ok(host);
  old.error({ message: "late failure" });
  host.complete();
  assert.equal((await replacement).ok, true);
  assert.equal(isPyodideLoaded(), true);
});
