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
      void api.agent.raw;
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

void test("createExtensionAPI registerTool fails fast when execute is missing", () => {
  let registerCalls = 0;

  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    registerTool: () => {
      registerCalls += 1;
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

  const badTool = {
    description: "Echo",
    parameters: Type.Object({
      text: Type.String(),
    }),
    handler: () => ({
      content: [{ type: "text", text: "Echo:hello" }],
      details: undefined,
    }),
  };

  assert.throws(
    () => {
      Reflect.apply(api.registerTool, api, ["echo", badTool]);
    },
    /use execute\(params, signal\?, onUpdate\?\) instead of handler/i,
  );

  assert.equal(registerCalls, 0);
});

void test("createExtensionAPI denies llm.complete when capability is blocked", async () => {
  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    llmComplete: () => Promise.resolve({
      content: "ok",
      model: "anthropic/claude-opus-4-6",
    }),
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "tools.register",
      "agent.read",
      "agent.events.read",
      "ui.overlay",
      "ui.widget",
      "ui.toast",
    ])),
    formatCapabilityError: (capability) => `DENIED:${capability}`,
  });

  await assert.rejects(
    async () => api.llm.complete({
      messages: [{ role: "user", content: "Hello" }],
    }),
    /DENIED:llm\.complete/,
  );
});

void test("createExtensionAPI forwards dynamic tools, storage, and agent steering", async () => {
  let unregisteredToolName = "";
  let injectedContext = "";
  let steeredMessage = "";
  let followUpMessage = "";

  const storage = new Map<string, unknown>();

  const api = createExtensionAPI({
    getAgent: () => {
      throw new Error("getAgent should not be called");
    },
    unregisterTool: (name) => {
      unregisteredToolName = name;
    },
    injectAgentContext: (content) => {
      injectedContext = content;
    },
    steerAgent: (content) => {
      steeredMessage = content;
    },
    followUpAgent: (content) => {
      followUpMessage = content;
    },
    storageGet: (key) => Promise.resolve(storage.get(key)),
    storageSet: (key, value) => {
      storage.set(key, value);
      return Promise.resolve();
    },
    storageDelete: (key) => {
      storage.delete(key);
      return Promise.resolve();
    },
    storageKeys: () => Promise.resolve(Array.from(storage.keys())),
    isCapabilityEnabled: createCapabilityGate(new Set<ExtensionCapability>([
      "commands.register",
      "tools.register",
      "agent.context.write",
      "agent.steer",
      "agent.followup",
      "storage.readwrite",
      "ui.overlay",
      "ui.widget",
      "ui.toast",
    ])),
  });

  api.unregisterTool("my-tool");
  api.agent.injectContext("context please");
  api.agent.steer("stop and rethink");
  api.agent.followUp("run review pass");

  await api.storage.set("alpha", 42);
  assert.equal(await api.storage.get("alpha"), 42);
  assert.deepEqual(await api.storage.keys(), ["alpha"]);
  await api.storage.delete("alpha");
  assert.deepEqual(await api.storage.keys(), []);

  assert.equal(unregisteredToolName, "my-tool");
  assert.equal(injectedContext, "context please");
  assert.equal(steeredMessage, "stop and rethink");
  assert.equal(followUpMessage, "run review pass");
});
