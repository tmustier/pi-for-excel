import assert from "node:assert/strict";
import { test } from "node:test";

import {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
} from "../src/tools/experimental-tool-gates.ts";
import type {
  LibreOfficeBridgeDetails,
  PythonBridgeDetails,
  TmuxBridgeDetails,
} from "../src/tools/tool-details.ts";
import {
  PYTHON_BRIDGE_SETUP_COMMAND,
  TMUX_BRIDGE_SETUP_COMMAND,
  resolveBridgeSetupCardModel,
  shouldShowBridgeSetupCard,
  testBridgeSetupConnection,
} from "../src/ui/bridge-setup-card.ts";

void test("shows tmux setup card for tmux bridge gate failures", () => {
  const details: TmuxBridgeDetails = {
    kind: "tmux_bridge",
    ok: false,
    action: "capture_pane",
    error: "Terminal access is not available right now because the tmux bridge is not reachable at the configured URL.",
    gateReason: "bridge_unreachable",
    skillHint: "tmux-bridge",
  };

  assert.equal(shouldShowBridgeSetupCard(details), true);

  const model = resolveBridgeSetupCardModel(details);
  assert.ok(model);
  assert.equal(model.command, TMUX_BRIDGE_SETUP_COMMAND);
  assert.equal(model.probeUrl, DEFAULT_TMUX_BRIDGE_URL);
  assert.match(model.title, /terminal access.*unavailable/i);
});

void test("uses bridge URL from details when testing tmux setup", () => {
  const details: TmuxBridgeDetails = {
    kind: "tmux_bridge",
    ok: false,
    action: "list_sessions",
    bridgeUrl: "https://localhost:4441",
    error: "Terminal access is not available right now because the tmux bridge is not reachable at the configured URL.",
    gateReason: "bridge_unreachable",
    skillHint: "tmux-bridge",
  };

  const model = resolveBridgeSetupCardModel(details);
  assert.ok(model);
  assert.equal(model.probeUrl, "https://localhost:4441");
});

void test("does not probe default URL when tmux bridge setting is invalid", async () => {
  const details: TmuxBridgeDetails = {
    kind: "tmux_bridge",
    ok: false,
    action: "list_sessions",
    error: "Terminal access is not available because the tmux bridge URL is invalid. Use a full URL like https://localhost:3341.",
    gateReason: "invalid_bridge_url",
    skillHint: "tmux-bridge",
  };

  const model = resolveBridgeSetupCardModel(details);
  assert.ok(model);
  assert.equal(model.probeUrl, null);

  let probeCalled = false;
  const reachable = await testBridgeSetupConnection(details, () => {
    probeCalled = true;
    return Promise.resolve(true);
  });

  assert.equal(reachable, false);
  assert.equal(probeCalled, false);
});

void test("shows python setup card when no Python runtime is available", () => {
  const details: PythonBridgeDetails = {
    kind: "python_bridge",
    ok: false,
    action: "run_python",
    error: "no_python_runtime",
    skillHint: "python-bridge",
  };

  assert.equal(shouldShowBridgeSetupCard(details), true);

  const model = resolveBridgeSetupCardModel(details);
  assert.ok(model);
  assert.equal(model.command, PYTHON_BRIDGE_SETUP_COMMAND);
  assert.equal(model.probeUrl, DEFAULT_PYTHON_BRIDGE_URL);
  assert.match(model.title, /python.*unavailable/i);
});

void test("shows python setup card for libreoffice bridge outages", () => {
  const details: LibreOfficeBridgeDetails = {
    kind: "libreoffice_bridge",
    ok: false,
    action: "convert",
    bridgeUrl: "https://localhost:4450",
    error: "Native Python is not available right now because the Python bridge is not reachable at the configured URL.",
    gateReason: "bridge_unreachable",
    skillHint: "python-bridge",
  };

  const model = resolveBridgeSetupCardModel(details);
  assert.ok(model);
  assert.equal(model.command, PYTHON_BRIDGE_SETUP_COMMAND);
  assert.equal(model.probeUrl, "https://localhost:4450");
  assert.equal(model.title, "File conversion is unavailable");
});

void test("does not show setup card for non-setup python errors", () => {
  const details: PythonBridgeDetails = {
    kind: "python_bridge",
    ok: false,
    action: "run_python",
    error: "NameError: x is not defined",
  };

  assert.equal(shouldShowBridgeSetupCard(details), false);
  assert.equal(resolveBridgeSetupCardModel(details), null);
});

void test("testBridgeSetupConnection probes the resolved URL", async () => {
  const details: PythonBridgeDetails = {
    kind: "python_bridge",
    ok: false,
    action: "run_python",
    bridgeUrl: "https://localhost:5540",
    error: "Native Python is not available right now because the Python bridge is not reachable at the configured URL.",
    gateReason: "bridge_unreachable",
    skillHint: "python-bridge",
  };

  let probedUrl: string | null = null;

  const reachable = await testBridgeSetupConnection(details, (url) => {
    probedUrl = url;
    return Promise.resolve(true);
  });

  assert.equal(reachable, true);
  assert.equal(probedUrl, "https://localhost:5540");
});
