import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";

import {
  activateExtensionInSandbox,
  type SandboxActivationOptions,
} from "../src/extensions/sandbox-runtime.ts";
import {
  isSandboxEnvelope,
  SANDBOX_BOOTSTRAP_KIND,
  SANDBOX_CHANNEL,
  type SandboxResponseEnvelope,
} from "../src/extensions/sandbox/protocol.ts";

interface SandboxHarness {
  port: MessagePort;
  close(): Promise<void>;
}

function defaultSandboxOptions(): SandboxActivationOptions {
  return {
    instanceId: "ext.sandbox.contract",
    extensionName: "Sandbox contract",
    source: { kind: "inline", code: "export function activate() {}" },
    registerCommand: () => {},
    registerTool: () => {},
    unregisterTool: () => {},
    subscribeAgentEvents: () => () => {},
    llmComplete: () => Promise.reject(new Error("not expected")),
    httpFetch: () => Promise.reject(new Error("not expected")),
    storageGet: () => Promise.resolve(null),
    storageSet: () => Promise.resolve(),
    storageDelete: () => Promise.resolve(),
    storageKeys: () => Promise.resolve([]),
    clipboardWriteText: () => Promise.resolve(),
    injectAgentContext: () => {},
    steerAgent: () => {},
    followUpAgent: () => {},
    listSkills: () => Promise.resolve([]),
    readSkill: () => Promise.reject(new Error("not expected")),
    installSkill: () => Promise.resolve(),
    uninstallSkill: () => Promise.resolve(),
    downloadFile: () => {},
    registerConnection: () => "connection-id",
    unregisterConnection: () => {},
    listConnections: () => Promise.resolve([]),
    getConnection: () => Promise.resolve(null),
    getConnectionSecrets: () => Promise.resolve(null),
    setConnectionSecrets: () => Promise.resolve(),
    clearConnectionSecrets: () => Promise.resolve(),
    markConnectionValidated: () => Promise.resolve(),
    markConnectionInvalid: () => Promise.resolve(),
    markConnectionStatus: () => Promise.resolve(),
    registerModelProvider: () => "provider-id",
    unregisterModelProvider: () => {},
    refreshModelProviders: () => Promise.resolve(),
    isCapabilityEnabled: () => true,
    formatCapabilityError: (capability) => `Denied ${capability}`,
    toast: () => {},
  };
}

async function openSandbox(
  overrides: Partial<SandboxActivationOptions> = {},
): Promise<SandboxHarness> {
  const previousDocument = Reflect.get(globalThis, "document");
  const previousWindow = Reflect.get(globalThis, "window");
  const previousHTMLElement = Reflect.get(globalThis, "HTMLElement");
  let iframe: HTMLElement | null = null;
  let sandboxPort: MessagePort | null = null;
  const getIframe = (): HTMLElement | null => iframe;
  const getSandboxPort = (): MessagePort | null => sandboxPort;

  const fakeDocument = {
    body: {
      appendChild: (element: HTMLElement): HTMLElement => {
        iframe = element;
        return element;
      },
    },
    createElement: (tagName: string): HTMLElement => {
      assert.equal(tagName, "iframe");
      const element = new EventTarget() as HTMLElement;
      const attributes = new Map<string, string>();
      Reflect.set(element, "style", {});
      Reflect.set(element, "setAttribute", (name: string, value: string) => attributes.set(name, value));
      Reflect.set(element, "getAttribute", (name: string) => attributes.get(name) ?? null);
      Reflect.set(element, "remove", () => {});
      Reflect.set(element, "contentWindow", {
        postMessage: (bootstrap: unknown, targetOrigin: string, transfer: MessagePort[]) => {
          assert.equal(targetOrigin, "*");
          assert.deepEqual(bootstrap, {
            channel: SANDBOX_CHANNEL,
            instanceId: overrides.instanceId ?? "ext.sandbox.contract",
            direction: "host_to_sandbox",
            kind: SANDBOX_BOOTSTRAP_KIND,
          });
          sandboxPort = transfer[0] ?? null;
        },
      });
      return element;
    },
    getElementById: () => null,
  };

  Reflect.set(globalThis, "document", fakeDocument);
  Reflect.set(globalThis, "window", {});
  Reflect.set(globalThis, "HTMLElement", EventTarget);

  const options = { ...defaultSandboxOptions(), ...overrides };
  const activation = activateExtensionInSandbox(options);
  const mountedIframe = getIframe();
  if (!mountedIframe) throw new Error("Expected sandbox iframe to mount.");
  mountedIframe.dispatchEvent(new Event("load"));

  const port = getSandboxPort();
  if (!port) throw new Error("Expected sandbox MessagePort transfer.");
  port.start();
  port.postMessage({
    channel: SANDBOX_CHANNEL,
    instanceId: options.instanceId,
    direction: "sandbox_to_host",
    kind: "event",
    event: "ready",
    data: null,
  });
  const handle = await activation;

  return {
    port,
    close: async () => {
      const onDeactivate = (event: MessageEvent<unknown>) => {
        const request = event.data;
        if (!isSandboxEnvelope(request) || request.kind !== "request" || request.method !== "deactivate") return;
        port.postMessage({
          channel: SANDBOX_CHANNEL,
          instanceId: options.instanceId,
          direction: "sandbox_to_host",
          kind: "response",
          requestId: request.requestId,
          ok: true,
          result: null,
        });
      };
      port.addEventListener("message", onDeactivate);
      await handle.deactivate();
      port.removeEventListener("message", onDeactivate);
      port.close();
      if (previousDocument === undefined) Reflect.deleteProperty(globalThis, "document");
      else Reflect.set(globalThis, "document", previousDocument);
      if (previousWindow === undefined) Reflect.deleteProperty(globalThis, "window");
      else Reflect.set(globalThis, "window", previousWindow);
      if (previousHTMLElement === undefined) Reflect.deleteProperty(globalThis, "HTMLElement");
      else Reflect.set(globalThis, "HTMLElement", previousHTMLElement);
    },
  };
}

function sendRequest(
  harness: SandboxHarness,
  method: string,
  params: unknown,
): Promise<SandboxResponseEnvelope> {
  const requestId = `${method}-request`;
  return new Promise((resolve) => {
    const onMessage = (event: MessageEvent<unknown>) => {
      const response = event.data;
      if (!isSandboxEnvelope(response) || response.kind !== "response" || response.requestId !== requestId) return;
      harness.port.removeEventListener("message", onMessage);
      resolve(response);
    };
    harness.port.addEventListener("message", onMessage);
    harness.port.postMessage({
      channel: SANDBOX_CHANNEL,
      instanceId: "ext.sandbox.contract",
      direction: "sandbox_to_host",
      kind: "request",
      requestId,
      method,
      params,
    });
  });
}

void test("a denied sandbox capability returns an error without invoking the host operation", async () => {
  let storageWrites = 0;
  const harness = await openSandbox({
    isCapabilityEnabled: (capability) => capability !== "storage.readwrite",
    storageSet: () => {
      storageWrites += 1;
      return Promise.resolve();
    },
  });

  try {
    const response = await sendRequest(harness, "storage_set", { key: "theme", value: "dark" });
    assert.equal(response.ok, false);
    assert.match(String(response.error), /Denied storage\.readwrite/u);
    assert.equal(storageWrites, 0);
  } finally {
    await harness.close();
  }
});

void test("sandbox capability denial takes precedence over malformed privileged payloads", async () => {
  const harness = await openSandbox({
    isCapabilityEnabled: (capability) => capability !== "ui.overlay",
  });

  try {
    const response = await sendRequest(harness, "overlay_show", null);
    assert.equal(response.ok, false);
    assert.match(String(response.error), /Denied ui\.overlay/u);
  } finally {
    await harness.close();
  }
});

void test("invalid sandbox requests return errors and do not reach host services", async () => {
  let storageReads = 0;
  const harness = await openSandbox({
    storageGet: () => {
      storageReads += 1;
      return Promise.resolve("value");
    },
  });

  try {
    const malformed = await sendRequest(harness, "storage_get", null);
    const unsupported = await sendRequest(harness, "launch_process", {});

    assert.equal(malformed.ok, false);
    assert.equal(unsupported.ok, false);
    assert.equal(storageReads, 0);
  } finally {
    await harness.close();
  }
});

void test("sandbox service failures are returned to the extension without breaking later requests", async () => {
  let attempts = 0;
  const harness = await openSandbox({
    storageGet: () => {
      attempts += 1;
      return attempts === 1
        ? Promise.reject(new Error("storage temporarily unavailable"))
        : Promise.resolve("recovered");
    },
  });

  try {
    const failed = await sendRequest(harness, "storage_get", { key: "theme" });
    const recovered = await sendRequest(harness, "storage_get", { key: "theme" });

    assert.equal(failed.ok, false);
    assert.equal(recovered.ok, true);
    assert.equal(recovered.result, "recovered");
  } finally {
    await harness.close();
  }
});

void test("reading connection secrets requires the dedicated secret capability", async () => {
  let secretReads = 0;
  const harness = await openSandbox({
    isCapabilityEnabled: (capability) => capability !== "connections.secrets.read",
    getConnectionSecrets: () => {
      secretReads += 1;
      return Promise.resolve({ token: "secret" });
    },
  });

  try {
    const response = await sendRequest(harness, "connections_get_secrets", {
      connectionId: "crm",
    });

    assert.equal(response.ok, false);
    assert.equal(secretReads, 0);
  } finally {
    await harness.close();
  }
});

void test("deactivating a sandbox extension releases its agent-event subscription", async () => {
  let unsubscribed = 0;
  const harness = await openSandbox({
    subscribeAgentEvents: () => () => {
      unsubscribed += 1;
    },
  });

  const response = await sendRequest(harness, "subscribe_agent_events", {
    subscriptionId: "updates",
  });
  assert.equal(response.ok, true);

  await harness.close();
  assert.equal(unsubscribed, 1);
});

void test("a sandbox extension can register and invoke a connection-qualified tool", async () => {
  const registrationState: { tool: AgentTool<TSchema, unknown> | null } = { tool: null };
  const harness = await openSandbox({
    registerTool: (tool) => {
      registrationState.tool = tool;
    },
  });

  try {
    const registration = await sendRequest(harness, "register_tool", {
      toolId: "lookup-tool",
      name: "lookup_company",
      description: "Looks up a company",
      parameters: { type: "object", properties: { company: { type: "string" } } },
      requiresConnection: ["crm", "ext.sandbox.contract.billing", "CRM"],
    });

    assert.equal(registration.ok, true);
    const tool = registrationState.tool;
    assert.ok(tool);
    assert.deepEqual(Reflect.get(tool, "requiresConnection"), [
      "ext.sandbox.contract.crm",
      "ext.sandbox.contract.billing",
    ]);

    const invocation = new Promise<void>((resolve) => {
      harness.port.addEventListener("message", (event: MessageEvent<unknown>) => {
        const request = event.data;
        if (!isSandboxEnvelope(request) || request.kind !== "request" || request.method !== "invoke_tool") return;
        harness.port.postMessage({
          channel: SANDBOX_CHANNEL,
          instanceId: "ext.sandbox.contract",
          direction: "sandbox_to_host",
          kind: "response",
          requestId: request.requestId,
          ok: true,
          result: { content: [{ type: "text", text: "Acme found" }] },
        });
        resolve();
      }, { once: true });
    });

    const result = await tool.execute("lookup-call", { company: "Acme" });
    await invocation;
    assert.deepEqual(result.content, [{ type: "text", text: "Acme found" }]);
  } finally {
    await harness.close();
  }
});
