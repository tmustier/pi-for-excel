import assert from "node:assert/strict";
import { test } from "node:test";

import {
  createExperimentalCommands,
  defaultProbeTmuxBridgeHealth,
} from "../src/commands/builtins/experimental.ts";
import { setExperimentalFeatureEnabled } from "../src/experiments/flags.ts";
import type {
  ExperimentalFeatureDefinition,
  ExperimentalFeatureId,
} from "../src/experiments/flags.ts";

const extensionPermissionsFeature: ExperimentalFeatureDefinition = {
  id: "extension_permission_gates",
  slug: "extension-permissions",
  aliases: ["extensions-permissions"],
  title: "Extension permission gates",
  description: "Enforce per-extension capability permissions when extensions activate.",
  wiring: "wired",
  storageKey: "pi.experimental.extensionPermissionGates",
};

function getExperimentalCommand(dependencies: Parameters<typeof createExperimentalCommands>[0]) {
  const command = createExperimentalCommands(dependencies).find((entry) => entry.name === "experimental");
  assert.ok(command);
  return command;
}

void test("experimental settings fail clearly when local storage is unavailable", () => {
  assert.throws(
    () => setExperimentalFeatureEnabled("ui_dark_mode", true),
    /Local settings storage is unavailable/u,
  );
});

void test("/experimental with no args opens the overlay", async () => {
  let openCount = 0;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {
      openCount += 1;
    },
    showToast: () => {},
  });

  await command.execute("");
  assert.equal(openCount, 1);
});

void test("/experimental help shows usage and feature list", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getFeatureSlugs: () => ["remote-extension-urls", "extension-permissions"],
  });

  await command.execute("help");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Usage:\s*\/experimental/u);
  assert.match(toasts[0], /extension-permissions/u);
  assert.match(toasts[0], /remote-extension-urls/u);
  assert.match(toasts[0], /tmux-bridge-url/u);
  assert.match(toasts[0], /tmux-bridge-token/u);
  assert.match(toasts[0], /tmux-status/u);
  assert.match(toasts[0], /python-bridge-url/u);
  assert.match(toasts[0], /python-bridge-token/u);
});

void test("/experimental on <feature> enables feature", async () => {
  const toasts: string[] = [];
  let setCall: [ExperimentalFeatureId, boolean] | null = null;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "extension-permissions" ? extensionPermissionsFeature : null),
    setFeatureEnabled: (featureId, enabled) => {
      setCall = [featureId, enabled];
    },
  });

  await command.execute("on extension-permissions");

  assert.deepEqual(setCall, ["extension_permission_gates", true]);
  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Extension permission gates:\s*enabled/u);
});

void test("/experimental off <feature> disables feature", async () => {
  const toasts: string[] = [];
  let setCall: [ExperimentalFeatureId, boolean] | null = null;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "extension-permissions" ? extensionPermissionsFeature : null),
    setFeatureEnabled: (featureId, enabled) => {
      setCall = [featureId, enabled];
    },
  });

  await command.execute("off extension-permissions");

  assert.deepEqual(setCall, ["extension_permission_gates", false]);
  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Extension permission gates:\s*disabled/u);
});

void test("/experimental toggle <feature> uses toggle result", async () => {
  const toasts: string[] = [];
  let toggledFeatureId: ExperimentalFeatureId | null = null;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "extension-permissions" ? extensionPermissionsFeature : null),
    toggleFeature: (featureId) => {
      toggledFeatureId = featureId;
      return true;
    },
  });

  await command.execute("toggle extension-permissions");

  assert.equal(toggledFeatureId, "extension_permission_gates");
  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Extension permission gates:\s*enabled/u);
});

void test("/experimental unknown action returns usage", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
  });

  await command.execute("wat tmux-bridge");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Usage:\s*\/experimental/u);
});

void test("/experimental on with unknown feature reports available slugs", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getFeatureSlugs: () => ["remote-extension-urls", "extension-permissions"],
    resolveFeature: () => null,
  });

  await command.execute("on does-not-exist");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Unknown feature:\s*does-not-exist/u);
  assert.match(toasts[0], /extension-permissions/u);
});

void test("/experimental on mcp-tools redirects to /tools alias", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: () => null,
  });

  await command.execute("on mcp-tools");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /managed in \/tools/u);
});

void test("/experimental tmux-bridge-url shows configured value", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3341"),
  });

  await command.execute("tmux-bridge-url");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge URL: https://localhost:3341");
});

void test("/experimental tmux-bridge-url <url> validates, stores URL, and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const stored: string[] = [];
  const changedConfigKeys: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    validateTmuxBridgeUrl: (url) => url.trim().replace(/\/+$/u, ""),
    setTmuxBridgeUrl: (url) => {
      stored.push(url);
      return Promise.resolve();
    },
    notifyToolConfigChanged: (configKey) => {
      changedConfigKeys.push(configKey);
    },
  });

  await command.execute("tmux-bridge-url https://localhost:3341/");

  assert.deepEqual(stored, ["https://localhost:3341"]);
  assert.deepEqual(changedConfigKeys, ["tmux.bridge.url"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge URL set to https://localhost:3341");
});

void test("/experimental tmux-bridge-url clear removes stored URL and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const changedConfigKeys: string[] = [];
  let clearCount = 0;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    clearTmuxBridgeUrl: () => {
      clearCount += 1;
      return Promise.resolve();
    },
    notifyToolConfigChanged: (configKey) => {
      changedConfigKeys.push(configKey);
    },
  });

  await command.execute("tmux-bridge-url clear");

  assert.equal(clearCount, 1);
  assert.deepEqual(changedConfigKeys, ["tmux.bridge.url"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge URL override cleared. Using default https://localhost:3341.");
});

void test("/experimental tmux-bridge-url invalid URL surfaces validation error", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    validateTmuxBridgeUrl: () => {
      throw new Error("Invalid Proxy URL");
    },
  });

  await command.execute("tmux-bridge-url ftp://localhost:3341");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Invalid Proxy URL");
});

void test("/experimental tmux-bridge-token shows masked configured value", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getTmuxBridgeToken: () => Promise.resolve("supersecrettoken"),
  });

  await command.execute("tmux-bridge-token");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge token: supe**********en (length 16)");
  assert.ok(!toasts[0].includes("supersecrettoken"));
});

void test("/experimental tmux-bridge-token <token> stores token and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const stored: string[] = [];
  const changedConfigKeys: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    validateTmuxBridgeToken: (token) => token.trim(),
    setTmuxBridgeToken: (token) => {
      stored.push(token);
      return Promise.resolve();
    },
    notifyToolConfigChanged: (configKey) => {
      changedConfigKeys.push(configKey);
    },
  });

  await command.execute("tmux-bridge-token supersecrettoken");

  assert.deepEqual(stored, ["supersecrettoken"]);
  assert.deepEqual(changedConfigKeys, ["tmux.bridge.token"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge token set (supe**********en).");
  assert.ok(!toasts[0].includes("supersecrettoken"));
});

void test("/experimental tmux-bridge-token clear removes token and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const changedConfigKeys: string[] = [];
  let clearCount = 0;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    clearTmuxBridgeToken: () => {
      clearCount += 1;
      return Promise.resolve();
    },
    notifyToolConfigChanged: (configKey) => {
      changedConfigKeys.push(configKey);
    },
  });

  await command.execute("tmux-bridge-token clear");

  assert.equal(clearCount, 1);
  assert.deepEqual(changedConfigKeys, ["tmux.bridge.token"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge token cleared.");
});

void test("/experimental tmux-bridge-token invalid token surfaces validation error", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    validateTmuxBridgeToken: () => {
      throw new Error("Tmux bridge token must not contain whitespace.");
    },
  });

  await command.execute("tmux-bridge-token has spaces");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge token must not contain whitespace.");
});

void test("/experimental tmux-status reports blocked state with hint", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getTmuxBridgeUrl: () => Promise.resolve(undefined),
    getTmuxBridgeToken: () => Promise.resolve(undefined),
    probeTmuxBridgeHealth: () => Promise.resolve({ reachable: false }),
  });

  await command.execute("tmux-status");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /bridge URL: https:\/\/localhost:3341 \(default\)/u);
  assert.match(toasts[0], /gate: blocked \(missing_bridge_url\)/u);
  assert.match(toasts[0], /health: unreachable/u);
});

void test("tmux status probe rejects HTTP success from a non-tmux backend", async () => {
  const originalFetch = globalThis.fetch;
  try {
    globalThis.fetch = () => Promise.resolve(new Response(JSON.stringify({
      ok: true,
      mode: "stub",
      backend: "stub",
    }), { status: 200 }));

    const status = await defaultProbeTmuxBridgeHealth("https://localhost:3341");
    assert.equal(status.reachable, false);
    assert.match(status.error ?? "", /did not identify a real tmux backend/u);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

void test("/experimental tmux-status reports healthy bridge details", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3341"),
    getTmuxBridgeToken: () => Promise.resolve("supersecrettoken"),
    probeTmuxBridgeHealth: () => Promise.resolve({
      reachable: true,
      status: 200,
      mode: "tmux",
      backend: "tmux",
      sessions: 2,
    }),
  });

  await command.execute("tmux-status");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /bridge URL: https:\/\/localhost:3341/u);
  assert.match(toasts[0], /auth token: set \(supe\*{10}en, length 16\)/u);
  assert.match(toasts[0], /gate: pass/u);
  assert.match(toasts[0], /health: reachable \(HTTP 200, mode=tmux, backend=tmux, sessions=2\)/u);
  assert.ok(!toasts[0].includes("supersecrettoken"));
});

void test("/experimental tmux-status uses a single health probe for gate + diagnostics", async () => {
  const toasts: string[] = [];
  let probeCount = 0;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3341"),
    getTmuxBridgeToken: () => Promise.resolve(undefined),
    probeTmuxBridgeHealth: () => {
      probeCount += 1;
      return Promise.resolve({
        reachable: false,
        status: 503,
        error: "bridge unavailable",
      });
    },
  });

  await command.execute("tmux-status");

  assert.equal(probeCount, 1);
  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /gate: blocked \(bridge_unreachable\)/u);
  assert.match(toasts[0], /health: unreachable \(HTTP 503; bridge unavailable\)/u);
});

void test("/experimental tmux-status with extra args shows usage", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
  });

  await command.execute("tmux-status extra");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Usage: /experimental tmux-status");
});

void test("/experimental python-bridge-url shows configured value", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getPythonBridgeUrl: () => Promise.resolve("https://localhost:3340"),
  });

  await command.execute("python-bridge-url");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Python bridge URL: https://localhost:3340");
});

void test("/experimental python-bridge-url <url> validates, stores URL, and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const stored: string[] = [];
  const changedConfigKeys: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    validatePythonBridgeUrl: (url) => url.trim().replace(/\/+$/u, ""),
    setPythonBridgeUrl: (url) => {
      stored.push(url);
      return Promise.resolve();
    },
    notifyToolConfigChanged: (configKey) => {
      changedConfigKeys.push(configKey);
    },
  });

  await command.execute("python-bridge-url https://localhost:3340/");

  assert.deepEqual(stored, ["https://localhost:3340"]);
  assert.deepEqual(changedConfigKeys, ["python.bridge.url"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Python bridge URL set to https://localhost:3340");
});

void test("/experimental python-bridge-url clear removes stored URL and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const changedConfigKeys: string[] = [];
  let clearCount = 0;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    clearPythonBridgeUrl: () => {
      clearCount += 1;
      return Promise.resolve();
    },
    notifyToolConfigChanged: (configKey) => {
      changedConfigKeys.push(configKey);
    },
  });

  await command.execute("python-bridge-url clear");

  assert.equal(clearCount, 1);
  assert.deepEqual(changedConfigKeys, ["python.bridge.url"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Python bridge URL override cleared. Using default https://localhost:3340.");
});

void test("/experimental python-bridge-token shows masked configured value", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getPythonBridgeToken: () => Promise.resolve("anothersecrettoken"),
  });

  await command.execute("python-bridge-token");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Python bridge token: anot************en (length 18)");
  assert.ok(!toasts[0].includes("anothersecrettoken"));
});

void test("/experimental python-bridge-token <token> stores token and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const stored: string[] = [];
  const changedConfigKeys: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    validatePythonBridgeToken: (token) => token.trim(),
    setPythonBridgeToken: (token) => {
      stored.push(token);
      return Promise.resolve();
    },
    notifyToolConfigChanged: (configKey) => {
      changedConfigKeys.push(configKey);
    },
  });

  await command.execute("python-bridge-token anothersecrettoken");

  assert.deepEqual(stored, ["anothersecrettoken"]);
  assert.deepEqual(changedConfigKeys, ["python.bridge.token"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Python bridge token set (anot************en).");
  assert.ok(!toasts[0].includes("anothersecrettoken"));
});

void test("/experimental python-bridge-token clear removes token and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const changedConfigKeys: string[] = [];
  let clearCount = 0;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    clearPythonBridgeToken: () => {
      clearCount += 1;
      return Promise.resolve();
    },
    notifyToolConfigChanged: (configKey) => {
      changedConfigKeys.push(configKey);
    },
  });

  await command.execute("python-bridge-token clear");

  assert.equal(clearCount, 1);
  assert.deepEqual(changedConfigKeys, ["python.bridge.token"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Python bridge token cleared.");
});

void test("/experimental python-bridge-token invalid token surfaces validation error", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    validatePythonBridgeToken: () => {
      throw new Error("Python bridge token must not contain whitespace.");
    },
  });

  await command.execute("python-bridge-token has spaces");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Python bridge token must not contain whitespace.");
});
