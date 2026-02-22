import assert from "node:assert/strict";
import { test } from "node:test";

import {
  probeLocalServices,
  type BridgeHealthDependencies,
  type LocalServiceEntry,
} from "../src/tools/bridge-health.ts";

function findPython(entries: LocalServiceEntry[]) {
  return entries.find((e) => e.name === "python");
}

function findTmux(entries: LocalServiceEntry[]) {
  return entries.find((e) => e.name === "tmux");
}

function makeDeps(overrides: Partial<BridgeHealthDependencies> = {}): BridgeHealthDependencies {
  return {
    getPythonBridgeUrl: () => Promise.resolve(undefined),
    getTmuxBridgeUrl: () => Promise.resolve(undefined),
    fetchHealth: () => Promise.resolve(null),
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// Both unreachable
// ---------------------------------------------------------------------------

void test("returns not_running for both when bridges are unreachable", async () => {
  const entries = await probeLocalServices(makeDeps());
  const python = findPython(entries);
  const tmux = findTmux(entries);

  assert.ok(python);
  assert.ok(tmux);
  assert.equal(python.status, "not_running");
  assert.equal(tmux.status, "not_running");
  assert.equal(python.skillName, "python-bridge");
  assert.equal(tmux.skillName, "tmux-bridge");
});

// ---------------------------------------------------------------------------
// Python bridge healthy with libreoffice
// ---------------------------------------------------------------------------

void test("parses fully healthy python bridge", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: (url) => {
      if (url.includes("3340")) {
        return Promise.resolve({
          ok: true,
          mode: "real",
          backend: "real",
          python: { available: true, command: "python3", version: "3.12.1" },
          libreoffice: { available: true, command: "soffice", version: "7.6.4" },
        });
      }
      return Promise.resolve(null);
    },
  }));

  const python = findPython(entries);
  assert.ok(python);
  assert.equal(python.status, "running");
  assert.equal(python.name === "python" && python.pythonVersion, "3.12.1");
  assert.equal(python.name === "python" && python.libreofficeAvailable, true);
  assert.equal(python.name === "python" && python.libreofficeVersion, "7.6.4");
});

void test("normalizes verbose libreoffice version strings", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: (url) => {
      if (url.includes("3340")) {
        return Promise.resolve({
          ok: true,
          mode: "real",
          python: { available: true, version: "3.12.1" },
          libreoffice: { available: true, version: "LibreOffice 7.6.4.1 40(Build:1)" },
        });
      }
      return Promise.resolve(null);
    },
  }));

  const python = findPython(entries);
  assert.ok(python);
  assert.equal(python.name === "python" && python.libreofficeVersion, "7.6.4.1");
});

void test("ignores non-numeric libreoffice version strings", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: (url) => {
      if (url.includes("3340")) {
        return Promise.resolve({
          ok: true,
          mode: "real",
          python: { available: true, version: "3.12.1" },
          libreoffice: { available: true, version: "ready and healthy" },
        });
      }
      return Promise.resolve(null);
    },
  }));

  const python = findPython(entries);
  assert.ok(python);
  assert.equal(python.name === "python" && python.libreofficeVersion, undefined);
});

// ---------------------------------------------------------------------------
// Python bridge up, libreoffice missing → partial
// ---------------------------------------------------------------------------

void test("reports partial when python available but libreoffice missing", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: (url) => {
      if (url.includes("3340")) {
        return Promise.resolve({
          ok: true,
          mode: "real",
          python: { available: true, command: "python3", version: "3.11.0" },
          libreoffice: { available: false, error: "No LibreOffice binary found" },
        });
      }
      return Promise.resolve(null);
    },
  }));

  const python = findPython(entries);
  assert.ok(python);
  assert.equal(python.status, "partial");
  assert.equal(python.name === "python" && python.pythonVersion, "3.11.0");
  assert.equal(python.name === "python" && python.libreofficeAvailable, false);
});

// ---------------------------------------------------------------------------
// Python bridge up, but Python binary missing → not_running
// ---------------------------------------------------------------------------

void test("reports not_running when bridge is up but python binary missing", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: (url) => {
      if (url.includes("3340")) {
        return Promise.resolve({
          ok: true,
          mode: "real",
          python: { available: false, error: "No Python binary found" },
          libreoffice: { available: false },
        });
      }
      return Promise.resolve(null);
    },
  }));

  const python = findPython(entries);
  assert.ok(python);
  assert.equal(python.status, "not_running");
});

// ---------------------------------------------------------------------------
// Tmux bridge fully healthy
// ---------------------------------------------------------------------------

void test("parses fully healthy tmux bridge", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: (url) => {
      if (url.includes("3341")) {
        return Promise.resolve({
          ok: true,
          mode: "tmux",
          backend: "tmux",
          tmuxVersion: "3.4",
          sessions: 2,
        });
      }
      return Promise.resolve(null);
    },
  }));

  const tmux = findTmux(entries);
  assert.ok(tmux);
  assert.equal(tmux.status, "running");
  assert.equal(tmux.name === "tmux" && tmux.tmuxVersion, "3.4");
  assert.equal(tmux.name === "tmux" && tmux.tmuxSessions, 2);
});

// ---------------------------------------------------------------------------
// Tmux bridge in stub mode → partial
// ---------------------------------------------------------------------------

void test("reports partial for tmux bridge in stub mode", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: (url) => {
      if (url.includes("3341")) {
        return Promise.resolve({
          ok: true,
          mode: "stub",
          backend: "stub",
          tmuxVersion: undefined,
          sessions: 0,
        });
      }
      return Promise.resolve(null);
    },
  }));

  const tmux = findTmux(entries);
  assert.ok(tmux);
  assert.equal(tmux.status, "partial");
});

// ---------------------------------------------------------------------------
// Probes run in parallel
// ---------------------------------------------------------------------------

void test("probes python and tmux in parallel", async () => {
  const callOrder: string[] = [];

  const entries = await probeLocalServices(makeDeps({
    fetchHealth: async (url) => {
      if (url.includes("3340")) {
        callOrder.push("python-start");
        await new Promise((r) => setTimeout(r, 20));
        callOrder.push("python-end");
        return { ok: true, python: { available: true, version: "3.12" }, libreoffice: { available: true } };
      }
      callOrder.push("tmux-start");
      await new Promise((r) => setTimeout(r, 20));
      callOrder.push("tmux-end");
      return { ok: true, mode: "tmux", tmuxVersion: "3.4", sessions: 1 };
    },
  }));

  // Both started before either ended → parallel execution
  assert.ok(callOrder.indexOf("python-start") < callOrder.indexOf("python-end"));
  assert.ok(callOrder.indexOf("tmux-start") < callOrder.indexOf("tmux-end"));
  // Interleaved start order proves parallelism
  const startIndices = [callOrder.indexOf("python-start"), callOrder.indexOf("tmux-start")];
  const endIndices = [callOrder.indexOf("python-end"), callOrder.indexOf("tmux-end")];
  assert.ok(Math.max(...startIndices) < Math.min(...endIndices));

  assert.equal(entries.length, 2);
});

// ---------------------------------------------------------------------------
// Malformed health payload → not_running
// ---------------------------------------------------------------------------

void test("handles malformed health payloads gracefully", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: () => Promise.resolve({ ok: false }),
  }));

  assert.equal(findPython(entries)?.status, "not_running");
  assert.equal(findTmux(entries)?.status, "not_running");
});

void test("handles non-object health payloads", async () => {
  const entries = await probeLocalServices(makeDeps({
    fetchHealth: () => Promise.resolve("not json"),
  }));

  assert.equal(findPython(entries)?.status, "not_running");
  assert.equal(findTmux(entries)?.status, "not_running");
});

// ---------------------------------------------------------------------------
// DisplayName and skillName are always correct
// ---------------------------------------------------------------------------

void test("entries always include correct displayName and skillName", async () => {
  const entries = await probeLocalServices(makeDeps());

  const python = findPython(entries);
  const tmux = findTmux(entries);
  assert.ok(python);
  assert.ok(tmux);
  assert.equal(python.displayName, "Python (native)");
  assert.equal(python.skillName, "python-bridge");
  assert.equal(tmux.displayName, "Terminal (tmux)");
  assert.equal(tmux.skillName, "tmux-bridge");
});
