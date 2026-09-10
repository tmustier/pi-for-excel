import assert from "node:assert/strict";
import { test } from "node:test";

import {
  detectSpreadsheetHost,
  resetCurrentSpreadsheetHostForTests,
  resolveSpreadsheetHostForBoot,
} from "../src/host/index.ts";

type HostGlobals = {
  Office?: unknown;
  wps?: unknown;
  Application?: unknown;
};

async function withHostGlobals<T>(values: HostGlobals, run: () => Promise<T> | T): Promise<T> {
  const keys = ["Office", "wps", "Application"] as const;
  const previous = keys.map((key) => ({
    key,
    present: Reflect.has(globalThis, key),
    value: Reflect.get(globalThis, key) as unknown,
  }));

  for (const key of keys) Reflect.deleteProperty(globalThis, key);
  for (const [key, value] of Object.entries(values)) Reflect.set(globalThis, key, value);

  try {
    return await run();
  } finally {
    for (const item of previous) {
      if (item.present) Reflect.set(globalThis, item.key, item.value);
      else Reflect.deleteProperty(globalThis, item.key);
    }
    resetCurrentSpreadsheetHostForTests();
  }
}

void test("POLICY: host globals and Office readiness select the boot host", async () => {
  const readyInfo = { host: "Excel", platform: "PC" };
  const readyOffice = {
    onReady: (callback: (info: typeof readyInfo) => void) => {
      callback(readyInfo);
      return Promise.resolve(readyInfo);
    },
  };
  const rows: ReadonlyArray<{
    globals: HostGlobals;
    detect: "wps" | "office" | "browser";
    boot?: { kind: "office" | "browser"; reason: "office-ready" | "office-timeout" | "office-unavailable" };
  }> = [
    { globals: { Office: readyOffice, wps: { EtApplication: () => ({}) } }, detect: "wps" },
    { globals: { wps: { Application: { ActiveWorkbook: null } } }, detect: "wps" },
    { globals: { wps: { PluginStorage: {} } }, detect: "wps" },
    { globals: { Application: { Range: () => ({}) } }, detect: "wps" },
    { globals: { Application: { CreateTaskpane: () => ({}) } }, detect: "wps" },
    // Unrelated globals with these names (e.g. injected by a browser extension) are not WPS.
    { globals: { Application: {} }, detect: "browser", boot: { kind: "browser", reason: "office-unavailable" } },
    { globals: { wps: {} }, detect: "browser", boot: { kind: "browser", reason: "office-unavailable" } },
    { globals: { Office: readyOffice, Application: { version: "1.0" } }, detect: "office", boot: { kind: "office", reason: "office-ready" } },
    { globals: { Office: readyOffice }, detect: "office", boot: { kind: "office", reason: "office-ready" } },
    { globals: { Office: { onReady: () => new Promise(() => {}) } }, detect: "office", boot: { kind: "office", reason: "office-timeout" } },
    { globals: { Office: { onReady: () => Promise.reject(new Error("not ready")) } }, detect: "office", boot: { kind: "office", reason: "office-timeout" } },
    { globals: {}, detect: "browser", boot: { kind: "browser", reason: "office-unavailable" } },
  ];

  for (const row of rows) {
    await withHostGlobals(row.globals, async () => {
      assert.equal(detectSpreadsheetHost(), row.detect);
      if (row.boot) {
        const result = await resolveSpreadsheetHostForBoot({ officeReadyTimeoutMs: 5 });
        assert.deepEqual(
          { kind: result.host.kind, reason: result.readyInfo.reason },
          row.boot,
        );
      }
    });
  }
});
