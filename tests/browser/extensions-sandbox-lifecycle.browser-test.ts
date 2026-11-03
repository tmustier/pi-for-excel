import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { openBrowserPage, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

void test("sandbox readiness includes registrations started while activation operations await", async () => {
  const { page, finish } = await openBrowserPage(env, { path: "src/ui-gallery.html" });
  try {
    const result = await page.evaluate(async () => {
      const srcdocPath = "/src/extensions/sandbox/srcdoc.ts";
      const protocolPath = "/src/extensions/sandbox/protocol.ts";
      const { buildSandboxSrcdoc } = await import(/* @vite-ignore */ srcdocPath) as typeof import("../../src/extensions/sandbox/srcdoc.ts");
      const { SANDBOX_CHANNEL, SANDBOX_BOOTSTRAP_KIND, isSandboxEnvelope } = await import(/* @vite-ignore */ protocolPath) as typeof import("../../src/extensions/sandbox/protocol.ts");
      const iframe = document.createElement("iframe");
      iframe.sandbox.add("allow-scripts");
      iframe.srcdoc = buildSandboxSrcdoc({ instanceId: "readiness", extensionName: "Readiness", widgetApiV2Enabled: false, source: { kind: "inline", code: `
        export function activate(api) {
          api.toast("initial");
          void api.storage.get("gate").then(() => api.registerCommand("late", { handler: () => {} }));
        }
      ` } });
      const channel = new MessageChannel();
      const events: string[] = [];
      let initialRequestId = "";
      let registrationId = "";
      let resolveProbe = (): void => {};
      let resolveReady = (): void => {};
      const probe = new Promise<void>((resolve) => { resolveProbe = resolve; });
      const ready = new Promise<void>((resolve) => { resolveReady = resolve; });
      const respond = (requestId: string) => channel.port1.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "readiness", direction: "host_to_sandbox", kind: "response", requestId, ok: true, result: null });
      channel.port1.onmessage = (event: MessageEvent<unknown>) => {
        const message = event.data;
        if (!isSandboxEnvelope(message)) return;
        if (message.kind === "event") {
          events.push(message.event);
          if (message.event === "ready") resolveReady();
        }
        if (message.kind === "response" && message.requestId === "probe") resolveProbe();
        if (message.kind !== "request") return;
        if (message.method === "toast") initialRequestId = message.requestId;
        if (message.method === "storage_get") respond(message.requestId);
        if (message.method === "register_command") {
          registrationId = message.requestId;
          const params = message.params as { commandId: string };
          respond(initialRequestId);
          channel.port1.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "readiness", direction: "host_to_sandbox", kind: "request", requestId: "probe", method: "invoke_command", params: { commandId: params.commandId, args: "" } });
        }
      };
      const loaded = new Promise<void>((resolve) => iframe.addEventListener("load", () => resolve(), { once: true }));
      document.body.appendChild(iframe);
      await loaded;
      iframe.contentWindow?.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "readiness", direction: "host_to_sandbox", kind: SANDBOX_BOOTSTRAP_KIND }, "*", [channel.port2]);
      await probe;
      const readyBeforeRegistration = events.includes("ready");
      respond(registrationId);
      await ready;
      iframe.remove();
      channel.port1.close();
      return { readyBeforeRegistration, events };
    });
    assert.equal(result.readyBeforeRegistration, false, "ready must wait for registrations added during the readiness barrier");
    assert.deepEqual(result.events, ["ready"]);
  } finally { await finish(); }
});

void test("sandbox rejects handler dispatch during and after deactivation", async () => {
  const { page, finish } = await openBrowserPage(env, { path: "src/ui-gallery.html" });
  try {
    const result = await page.evaluate(async () => {
      const srcdocPath = "/src/extensions/sandbox/srcdoc.ts";
      const protocolPath = "/src/extensions/sandbox/protocol.ts";
      const { buildSandboxSrcdoc } = await import(/* @vite-ignore */ srcdocPath) as typeof import("../../src/extensions/sandbox/srcdoc.ts");
      const { SANDBOX_CHANNEL, SANDBOX_BOOTSTRAP_KIND, isSandboxEnvelope } = await import(/* @vite-ignore */ protocolPath) as typeof import("../../src/extensions/sandbox/protocol.ts");
      const iframe = document.createElement("iframe");
      iframe.sandbox.add("allow-scripts");
      iframe.srcdoc = buildSandboxSrcdoc({ instanceId: "terminal", extensionName: "Terminal", widgetApiV2Enabled: false, source: { kind: "inline", code: `
        export function activate(api) {
          api.registerCommand("invoke", { handler: () => api.toast("invoked") });
          return () => api.storage.get("cleanup");
        }
      ` } });
      const channel = new MessageChannel();
      const pending = new Map<string, (ok: boolean) => void>();
      let commandId = "";
      let cleanupId = "";
      let invoked = 0;
      let resolveReady = (): void => {};
      let resolveCleanup = (): void => {};
      const ready = new Promise<void>((resolve) => { resolveReady = resolve; });
      const cleanup = new Promise<void>((resolve) => { resolveCleanup = resolve; });
      const respond = (requestId: string) => channel.port1.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "terminal", direction: "host_to_sandbox", kind: "response", requestId, ok: true, result: null });
      const request = (requestId: string, method: string) => new Promise<boolean>((resolve) => {
        pending.set(requestId, resolve);
        channel.port1.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "terminal", direction: "host_to_sandbox", kind: "request", requestId, method, params: { commandId, args: "" } });
      });
      channel.port1.onmessage = (event: MessageEvent<unknown>) => {
        const message = event.data;
        if (!isSandboxEnvelope(message)) return;
        if (message.kind === "event" && message.event === "ready") resolveReady();
        if (message.kind === "response") pending.get(message.requestId)?.(message.ok);
        if (message.kind !== "request") return;
        if (message.method === "register_command") commandId = (message.params as { commandId: string }).commandId;
        if (message.method === "storage_get") {
          cleanupId = message.requestId;
          resolveCleanup();
          return;
        }
        if (message.method === "toast") invoked += 1;
        respond(message.requestId);
      };
      const loaded = new Promise<void>((resolve) => iframe.addEventListener("load", () => resolve(), { once: true }));
      document.body.appendChild(iframe);
      await loaded;
      iframe.contentWindow?.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "terminal", direction: "host_to_sandbox", kind: SANDBOX_BOOTSTRAP_KIND }, "*", [channel.port2]);
      await ready;
      const stopping = request("stop", "deactivate");
      await cleanup;
      const during = await request("during", "invoke_command");
      respond(cleanupId);
      await stopping;
      const after = await request("after", "invoke_command");
      iframe.remove();
      channel.port1.close();
      return { during, after, invoked };
    });
    assert.equal(result.during, false, "deactivating sandbox must not invoke retained handlers");
    assert.equal(result.after, false, "disposed sandbox must not invoke retained handlers");
    assert.equal(result.invoked, 0);
  } finally { await finish(); }
});

void test("sandbox host RPCs time out instead of retaining command invocations forever", async () => {
  const { page, finish } = await openBrowserPage(env, { path: "src/ui-gallery.html" });
  try {
    await page.clock.install();
    const pending = page.evaluate(async () => {
      const srcdocPath = "/src/extensions/sandbox/srcdoc.ts";
      const protocolPath = "/src/extensions/sandbox/protocol.ts";
      const { buildSandboxSrcdoc } = await import(/* @vite-ignore */ srcdocPath) as typeof import("../../src/extensions/sandbox/srcdoc.ts");
      const { SANDBOX_CHANNEL, SANDBOX_BOOTSTRAP_KIND, isSandboxEnvelope } = await import(/* @vite-ignore */ protocolPath) as typeof import("../../src/extensions/sandbox/protocol.ts");
      const iframe = document.createElement("iframe");
      iframe.sandbox.add("allow-scripts");
      iframe.srcdoc = buildSandboxSrcdoc({ instanceId: "timeout", extensionName: "Timeout", widgetApiV2Enabled: false, source: { kind: "inline", code: `
        export function activate(api) {
          api.registerCommand("pending", { handler: () => api.storage.get("never-replied") });
        }
      ` } });
      const channel = new MessageChannel();
      let commandId = "";
      let resolveReady = (): void => {};
      let resolveProbe = (_result: string): void => {};
      const ready = new Promise<void>((resolve) => { resolveReady = resolve; });
      const probe = new Promise<string>((resolve) => { resolveProbe = resolve; });
      channel.port1.onmessage = (event: MessageEvent<unknown>) => {
        const message = event.data;
        if (!isSandboxEnvelope(message)) return;
        if (message.kind === "event" && message.event === "ready") resolveReady();
        if (message.kind === "response" && message.requestId === "invoke") resolveProbe(message.error ?? "");
        if (message.kind !== "request") return;
        if (message.method === "storage_get") {
          document.body.dataset.rpcWaiting = "true";
          return;
        }
        if (message.method === "register_command") commandId = (message.params as { commandId: string }).commandId;
        channel.port1.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "timeout", direction: "host_to_sandbox", kind: "response", requestId: message.requestId, ok: true, result: null });
      };
      const loaded = new Promise<void>((resolve) => iframe.addEventListener("load", () => resolve(), { once: true }));
      document.body.appendChild(iframe);
      await loaded;
      iframe.contentWindow?.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "timeout", direction: "host_to_sandbox", kind: SANDBOX_BOOTSTRAP_KIND }, "*", [channel.port2]);
      await ready;
      channel.port1.postMessage({ channel: SANDBOX_CHANNEL, instanceId: "timeout", direction: "host_to_sandbox", kind: "request", requestId: "invoke", method: "invoke_command", params: { commandId } });
      const error = await probe;
      iframe.remove();
      channel.port1.close();
      return error;
    });
    await page.waitForFunction(() => document.body.dataset.rpcWaiting === "true");
    await page.clock.runFor(30_001);
    const result = await Promise.race([pending, page.evaluate(() => "still pending")]);
    assert.match(result, /host request timed out: storage_get/, "unanswered host RPC must reject at the shared timeout");
  } finally { await finish(); }
});
