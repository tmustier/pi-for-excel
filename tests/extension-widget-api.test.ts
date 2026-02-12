import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";

import { createExtensionAPI } from "../src/commands/extension-api.ts";
import type { ExtensionCapability } from "../src/extensions/permissions.ts";

function createCapabilityGate(allowed: ReadonlySet<ExtensionCapability>) {
  return (capability: ExtensionCapability): boolean => allowed.has(capability);
}

void test("widget v2 methods throw clear guidance when experiment is disabled", () => {
  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    widgetApiV2Enabled: false,
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "tools.register",
      "agent.read",
      "agent.events.read",
      "ui.overlay",
      "ui.widget",
      "ui.toast",
    ])),
  });

  assert.throws(
    () => {
      api.widget.remove("summary");
    },
    /Widget API v2 is disabled\./,
  );

  assert.throws(
    () => {
      api.widget.clear();
    },
    /extension-widget-v2/,
  );
});

void test("widget v2 methods still enforce ui.widget capability when enabled", () => {
  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    widgetApiV2Enabled: true,
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "tools.register",
      "agent.read",
      "agent.events.read",
      "ui.overlay",
      "ui.toast",
    ])),
    formatCapabilityError: (capability) => `DENIED:${capability}`,
  });

  assert.throws(
    () => {
      api.widget.remove("summary");
    },
    /DENIED:ui\.widget/,
  );
});

void test("extension API source exports additive widget lifecycle methods", async () => {
  const source = await readFile(new URL("../src/commands/extension-api.ts", import.meta.url), "utf8");

  assert.match(source, /upsert\(spec: WidgetUpsertSpec\)/);
  assert.match(source, /remove\(id: string\)/);
  assert.match(source, /clear\(\)/);
  assert.match(source, /extension_widget_v2/);
});
