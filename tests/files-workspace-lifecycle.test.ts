import "fake-indexeddb/auto";
import assert from "node:assert/strict";
import { test, type TestContext } from "node:test";

import { MemoryBackend } from "../src/files/backend.ts";
import { FilesWorkspace } from "../src/files/workspace.ts";
import { initAppStorage } from "../src/storage/init-app-storage.ts";

const NATIVE_HANDLE_KEY = "files.workspace.nativeHandle.v1";

function directoryHandle(name: string) {
  return {
    kind: "directory" as const,
    name,
    queryPermission: () => Promise.resolve("granted"),
    requestPermission: () => Promise.resolve("granted"),
    getDirectoryHandle: () => Promise.reject(new Error("directory not found")),
    getFileHandle: () => Promise.reject(new Error("file not found")),
    async *entries() {},
  };
}

function installHost(t: TestContext, picker: () => Promise<ReturnType<typeof directoryHandle>>): void {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "window");
  Object.defineProperty(globalThis, "window", {
    configurable: true,
    value: { top: null, showDirectoryPicker: picker },
  });
  t.after(() => {
    if (previous) Object.defineProperty(globalThis, "window", previous);
    else Reflect.deleteProperty(globalThis, "window");
  });
}

function deferred<T = void>(): { promise: Promise<T>; resolve: (value: T) => void } {
  let resolve: (value: T) => void = () => {};
  const promise = new Promise<T>((fulfill) => { resolve = fulfill; });
  return { promise, resolve };
}

void test("disconnect waits for an earlier folder connection and preserves uploaded files", async (t) => {
  const picked = deferred<ReturnType<typeof directoryHandle>>();
  const opened = deferred<void>();
  installHost(t, () => {
    opened.resolve();
    return picked.promise;
  });
  const backend = new MemoryBackend();
  await backend.writeBytes("uploaded.txt", new TextEncoder().encode("preserved"));
  const workspace = new FilesWorkspace({ initialBackend: backend, settings: null });
  const connecting = workspace.connectNativeDirectory();
  await opened.promise;
  const disconnecting = workspace.disconnectNativeDirectory();
  picked.resolve(directoryHandle("Selected"));
  await Promise.all([connecting, disconnecting]);

  const status = await workspace.getBackendStatus();
  assert.equal(status.kind, "memory", "the later disconnect must win over the pending connection");
  assert.equal(status.nativeDirectoryName, undefined);
  assert.equal((await workspace.readFile("uploaded.txt")).text, "preserved");
});

void test("disconnect during initialization cannot retain a restored folder name", async () => {
  const { settings } = initAppStorage(`workspace-lifecycle-${crypto.randomUUID()}`);
  const reading = deferred<void>();
  const restored = deferred<ReturnType<typeof directoryHandle>>();
  settings.get = (key) => {
    if (key !== NATIVE_HANDLE_KEY) return Promise.resolve(null);
    reading.resolve();
    return restored.promise;
  };
  settings.delete = () => Promise.resolve();
  const workspace = new FilesWorkspace({ settings: null });
  const statusDuringInitialization = workspace.getBackendStatus();
  await reading.promise;
  const disconnecting = workspace.disconnectNativeDirectory();
  restored.resolve(directoryHandle("Restored"));
  await Promise.all([statusDuringInitialization, disconnecting]);

  const status = await workspace.getBackendStatus();
  assert.equal(status.nativeConnected, false);
  assert.equal(status.nativeDirectoryName, undefined, "a disconnected backend must not report the restored folder");
});

void test("backend initialization can be retried after a host failure", async (t) => {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "navigator");
  let attempts = 0;
  Object.defineProperty(globalThis, "navigator", {
    configurable: true,
    value: {
      get storage() {
        attempts += 1;
        if (attempts === 1) throw new Error("temporary storage access failure");
        return undefined;
      },
    },
  });
  t.after(() => {
    if (previous) Object.defineProperty(globalThis, "navigator", previous);
    else Reflect.deleteProperty(globalThis, "navigator");
  });
  const { settings } = initAppStorage(`workspace-retry-${crypto.randomUUID()}`);
  settings.get = () => Promise.resolve(null);
  const workspace = new FilesWorkspace({ settings: null });
  await assert.rejects(workspace.getBackendStatus(), /temporary storage access failure/);
  assert.equal((await workspace.getBackendStatus()).kind, "memory");
  await workspace.writeTextFile("retried.txt", "ready");
  assert.equal((await workspace.readFile("retried.txt")).text, "ready");
});
