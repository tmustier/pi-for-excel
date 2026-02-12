import assert from "node:assert/strict";
import { test } from "node:test";

import { createExperimentalCommands } from "../src/commands/builtins/experimental.ts";
import type {
  ExperimentalFeatureDefinition,
  ExperimentalFeatureId,
} from "../src/experiments/flags.ts";

const tmuxFeature: ExperimentalFeatureDefinition = {
  id: "tmux_bridge",
  slug: "tmux-bridge",
  aliases: ["tmux"],
  title: "Tmux local bridge",
  description: "Allow local tmux bridge integration.",
  wiring: "flag-only",
  storageKey: "pi.experimental.tmuxBridge",
};

function getExperimentalCommand(dependencies: Parameters<typeof createExperimentalCommands>[0]) {
  const command = createExperimentalCommands(dependencies).find((entry) => entry.name === "experimental");
  assert.ok(command);
  return command;
}

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
    getFeatureSlugs: () => ["tmux-bridge", "remote-extension-urls"],
  });

  await command.execute("help");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Usage:\s*\/experimental/u);
  assert.match(toasts[0], /tmux-bridge/u);
  assert.match(toasts[0], /remote-extension-urls/u);
  assert.match(toasts[0], /tmux-bridge-url/u);
  assert.match(toasts[0], /tmux-bridge-token/u);
  assert.match(toasts[0], /tmux-status/u);
  assert.match(toasts[0], /python-bridge-url/u);
  assert.match(toasts[0], /python-bridge-token/u);
});

void test("/experimental on <feature> enables feature and reports flag-only suffix", async () => {
  const toasts: string[] = [];
  let setCall: [ExperimentalFeatureId, boolean] | null = null;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "tmux-bridge" ? tmuxFeature : null),
    setFeatureEnabled: (featureId, enabled) => {
      setCall = [featureId, enabled];
    },
  });

  await command.execute("on tmux-bridge");

  assert.deepEqual(setCall, ["tmux_bridge", true]);
  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Tmux local bridge:\s*enabled/u);
  assert.match(toasts[0], /flag saved; feature not wired yet/u);
});

void test("/experimental off <feature> disables feature", async () => {
  const toasts: string[] = [];
  let setCall: [ExperimentalFeatureId, boolean] | null = null;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "tmux-bridge" ? tmuxFeature : null),
    setFeatureEnabled: (featureId, enabled) => {
      setCall = [featureId, enabled];
    },
  });

  await command.execute("off tmux-bridge");

  assert.deepEqual(setCall, ["tmux_bridge", false]);
  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Tmux local bridge:\s*disabled/u);
});

void test("/experimental toggle <feature> uses toggle result", async () => {
  const toasts: string[] = [];
  let toggledFeatureId: ExperimentalFeatureId | null = null;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "tmux-bridge" ? tmuxFeature : null),
    toggleFeature: (featureId) => {
      toggledFeatureId = featureId;
      return true;
    },
  });

  await command.execute("toggle tmux-bridge");

  assert.equal(toggledFeatureId, "tmux_bridge");
  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Tmux local bridge:\s*enabled/u);
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
    getFeatureSlugs: () => ["tmux-bridge", "remote-extension-urls"],
    resolveFeature: () => null,
  });

  await command.execute("on does-not-exist");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /Unknown feature:\s*does-not-exist/u);
  assert.match(toasts[0], /tmux-bridge/u);
});

void test("/experimental tmux-bridge-url shows configured value", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3337"),
  });

  await command.execute("tmux-bridge-url");

  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge URL: https://localhost:3337");
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

  await command.execute("tmux-bridge-url https://localhost:3337/");

  assert.deepEqual(stored, ["https://localhost:3337"]);
  assert.deepEqual(changedConfigKeys, ["tmux.bridge.url"]);
  assert.equal(toasts.length, 1);
  assert.equal(toasts[0], "Tmux bridge URL set to https://localhost:3337");
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
  assert.equal(toasts[0], "Tmux bridge URL cleared.");
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

  await command.execute("tmux-bridge-url ftp://localhost:3337");

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
    isTmuxBridgeEnabled: () => false,
    getTmuxBridgeUrl: () => Promise.resolve(undefined),
    getTmuxBridgeToken: () => Promise.resolve(undefined),
  });

  await command.execute("tmux-status");

  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /feature flag \(tmux-bridge\): disabled/u);
  assert.match(toasts[0], /gate: blocked \(tmux_experiment_disabled\)/u);
  assert.match(toasts[0], /Enable it with \/experimental on tmux-bridge/u);
  assert.match(toasts[0], /health: not checked/u);
});

void test("/experimental tmux-status reports healthy bridge details", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    isTmuxBridgeEnabled: () => true,
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3337"),
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
  assert.match(toasts[0], /feature flag \(tmux-bridge\): enabled/u);
  assert.match(toasts[0], /bridge URL: https:\/\/localhost:3337/u);
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
    isTmuxBridgeEnabled: () => true,
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3337"),
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
  assert.equal(toasts[0], "Python bridge URL cleared.");
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
