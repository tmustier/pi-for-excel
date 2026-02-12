import assert from "node:assert/strict";
import { test } from "node:test";

import { Type } from "@sinclair/typebox";

import { createExtensionAPI } from "../src/commands/extension-api.ts";
import type { ExtensionCapability } from "../src/extensions/permissions.ts";

function createCapabilityGate(allowed: ReadonlySet<ExtensionCapability>) {
  return (capability: ExtensionCapability): boolean => allowed.has(capability);
}

void test("createExtensionAPI denies registerCommand when capability is blocked", () => {
  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    registerCommand: () => {
      throw new Error("registerCommand should not be called");
    },
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "tools.register",
      "agent.read",
      "agent.events.read",
      "ui.overlay",
      "ui.widget",
      "ui.toast",
    ])),
    formatCapabilityError: (capability) => `DENIED:${capability}`,
  });

  assert.throws(
    () => {
      api.registerCommand("hello", {
        description: "Hello",
        handler: () => {},
      });
    },
    /DENIED:commands\.register/,
  );
});

void test("createExtensionAPI denies registerTool when capability is blocked", () => {
  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    registerTool: () => {
      throw new Error("registerTool should not be called");
    },
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "agent.read",
      "agent.events.read",
      "ui.overlay",
      "ui.widget",
      "ui.toast",
    ])),
    formatCapabilityError: (capability) => `DENIED:${capability}`,
  });

  assert.throws(
    () => {
      api.registerTool("echo", {
        description: "Echo",
        parameters: Type.Object({
          text: Type.String(),
        }),
        execute: () => ({
          content: [{ type: "text", text: "ok" }],
          details: undefined,
        }),
      });
    },
    /DENIED:tools\.register/,
  );
});

void test("createExtensionAPI denies raw agent access before getAgent call", () => {
  let getAgentCalls = 0;

  const api = createExtensionAPI({
    getAgent: () => {
      getAgentCalls += 1;
      throw new Error("getAgent should not be called");
    },
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "tools.register",
      "agent.events.read",
      "ui.overlay",
      "ui.widget",
      "ui.toast",
    ])),
    formatCapabilityError: (capability) => `DENIED:${capability}`,
  });

  assert.throws(
    () => {
      void api.agent;
    },
    /DENIED:agent\.read/,
  );

  assert.equal(getAgentCalls, 0);
});

void test("createExtensionAPI denies onAgentEvent when agent.events.read is blocked", () => {
  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    subscribeAgentEvents: () => {
      throw new Error("subscribeAgentEvents should not be called");
    },
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "tools.register",
      "agent.read",
      "ui.overlay",
      "ui.widget",
      "ui.toast",
    ])),
    formatCapabilityError: (capability) => `DENIED:${capability}`,
  });

  assert.throws(
    () => {
      api.onAgentEvent(() => {});
    },
    /DENIED:agent\.events\.read/,
  );
});

void test("createExtensionAPI registerTool forwards metadata to host registrar", () => {
  let registeredName = "";
  let registeredLabel = "";
  let registeredDescription = "";

  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    registerTool: (tool) => {
      registeredName = tool.name;
      registeredLabel = tool.label;
      registeredDescription = tool.description;
    },
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

  api.registerTool("echo", {
    description: "Echo",
    parameters: Type.Object({
      text: Type.String(),
    }),
    execute: () => ({
      content: [{ type: "text", text: "Echo:hello" }],
      details: { len: 5 },
    }),
  });

  assert.equal(registeredName, "echo");
  assert.equal(registeredLabel, "echo");
  assert.equal(registeredDescription, "Echo");
});
