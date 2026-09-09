import assert from "node:assert/strict";
import { test } from "node:test";

import { createExperimentalCommands } from "../src/commands/builtins/experimental.ts";
import type {
  ExperimentalFeatureDefinition,
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

void test("/experimental with no args opens the overlay", async () => {
  let overlayOpened = false;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {
      overlayOpened = true;
    },
    showToast: () => {},
  });

  await command.execute("");
  assert.equal(overlayOpened, true);
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
  let enabled = false;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "extension-permissions" ? extensionPermissionsFeature : null),
    setFeatureEnabled: (featureId, value) => {
      assert.equal(featureId, "extension_permission_gates");
      enabled = value;
    },
  });

  await command.execute("on extension-permissions");

  assert.equal(enabled, true);
  assert.match(toasts[0], /Extension permission gates:\s*enabled/u);
});

void test("/experimental off <feature> disables feature", async () => {
  const toasts: string[] = [];
  let enabled = true;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "extension-permissions" ? extensionPermissionsFeature : null),
    setFeatureEnabled: (featureId, value) => {
      assert.equal(featureId, "extension_permission_gates");
      enabled = value;
    },
  });

  await command.execute("off extension-permissions");

  assert.equal(enabled, false);
  assert.match(toasts[0], /Extension permission gates:\s*disabled/u);
});

void test("/experimental toggle <feature> uses toggle result", async () => {
  const toasts: string[] = [];
  let enabled = false;

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    resolveFeature: (input) => (input === "extension-permissions" ? extensionPermissionsFeature : null),
    toggleFeature: (featureId) => {
      assert.equal(featureId, "extension_permission_gates");
      enabled = true;
      return enabled;
    },
  });

  await command.execute("toggle extension-permissions");

  assert.equal(enabled, true);
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

  assert.equal(toasts[0], "Tmux bridge URL: https://localhost:3341");
});

void test("/experimental tmux-bridge-url <url> validates, stores URL, and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const stored: string[] = [];

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
  });

  await command.execute("tmux-bridge-url https://localhost:3341/");

  assert.equal(stored.includes("https://localhost:3341"), true);
  assert.equal(toasts[0], "Tmux bridge URL set to https://localhost:3341");
});

void test("/experimental tmux-bridge-url clear removes stored URL and triggers tool refresh notice", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    clearTmuxBridgeUrl: () => Promise.resolve(),
  });

  await command.execute("tmux-bridge-url clear");

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

  assert.equal(toasts[0], "Tmux bridge token: supe**********en (length 16)");
  assert.ok(!toasts[0].includes("supersecrettoken"));
});

void test("/experimental tmux-bridge-token <token> stores token and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const stored: string[] = [];

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
  });

  await command.execute("tmux-bridge-token supersecrettoken");

  assert.equal(stored.includes("supersecrettoken"), true);
  assert.equal(toasts[0], "Tmux bridge token set (supe**********en).");
  assert.ok(!toasts[0].includes("supersecrettoken"));
});

void test("/experimental tmux-bridge-token clear removes token and triggers tool refresh notice", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    clearTmuxBridgeToken: () => Promise.resolve(),
  });

  await command.execute("tmux-bridge-token clear");

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
  });

  await command.execute("tmux-status");

  assert.match(toasts[0], /bridge URL: not set/u);
  assert.match(toasts[0], /gate: blocked \(missing_bridge_url\)/u);
  assert.match(toasts[0], /health: not checked/u);
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

  assert.match(toasts[0], /bridge URL: https:\/\/localhost:3341/u);
  assert.match(toasts[0], /auth token: set \(supe\*{10}en, length 16\)/u);
  assert.match(toasts[0], /gate: pass/u);
  assert.match(toasts[0], /health: reachable \(HTTP 200, mode=tmux, backend=tmux, sessions=2\)/u);
  assert.ok(!toasts[0].includes("supersecrettoken"));
});

void test("/experimental tmux-status uses a single health probe for gate + diagnostics", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    getTmuxBridgeUrl: () => Promise.resolve("https://localhost:3341"),
    getTmuxBridgeToken: () => Promise.resolve(undefined),
    probeTmuxBridgeHealth: () => {
      return Promise.resolve({
        reachable: false,
        status: 503,
        error: "bridge unavailable",
      });
    },
  });

  await command.execute("tmux-status");

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

  assert.equal(toasts[0], "Python bridge URL: https://localhost:3340");
});

void test("/experimental python-bridge-url <url> validates, stores URL, and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const stored: string[] = [];

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
  });

  await command.execute("python-bridge-url https://localhost:3340/");

  assert.equal(stored.includes("https://localhost:3340"), true);
  assert.equal(toasts[0], "Python bridge URL set to https://localhost:3340");
});

void test("/experimental python-bridge-url clear removes stored URL and triggers tool refresh notice", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    clearPythonBridgeUrl: () => Promise.resolve(),
  });

  await command.execute("python-bridge-url clear");

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

  assert.equal(toasts[0], "Python bridge token: anot************en (length 18)");
  assert.ok(!toasts[0].includes("anothersecrettoken"));
});

void test("/experimental python-bridge-token <token> stores token and triggers tool refresh notice", async () => {
  const toasts: string[] = [];
  const stored: string[] = [];

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
  });

  await command.execute("python-bridge-token anothersecrettoken");

  assert.equal(stored.includes("anothersecrettoken"), true);
  assert.equal(toasts[0], "Python bridge token set (anot************en).");
  assert.ok(!toasts[0].includes("anothersecrettoken"));
});

void test("/experimental python-bridge-token clear removes token and triggers tool refresh notice", async () => {
  const toasts: string[] = [];

  const command = getExperimentalCommand({
    showExperimentalDialog: () => {},
    showToast: (message) => {
      toasts.push(message);
    },
    clearPythonBridgeToken: () => Promise.resolve(),
  });

  await command.execute("python-bridge-token clear");

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

  assert.equal(toasts[0], "Python bridge token must not contain whitespace.");
});
