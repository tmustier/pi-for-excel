import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

const responseSchema = Type.Object({
  commandId: Type.String(),
  ok: Type.Boolean(),
  result: Type.Optional(Type.Unknown()),
  error: Type.Optional(Type.Object({ message: Type.String() })),
});
type BridgeResponse = Static<typeof responseSchema>;
const installedSchema = Type.Object({ extensionId: Type.String() });
const extensionsSchema = Type.Object({ extensions: Type.Array(Type.Object({
  name: Type.String(), loaded: Type.Boolean(), runtimeMode: Type.String(), lastError: Type.Union([Type.String(), Type.Null()]),
})) });
const counterResultSchema = Type.Object({
  content: Type.Array(Type.Object({ type: Type.Literal("text"), text: Type.String() })),
  details: Type.Object({ count: Type.Integer() }),
});
const storageResultSchema = Type.Object({ value: Type.Integer() });

// Real taskpane, module import, opaque iframe + MessagePort, capability refresh,
// Durable direct task and persisted extension storage. Only the bridge transport
// is intercepted; no model, credentials, native Excel or workbook effects.
void test("inline sandbox tools stay executable across install, pane reload, extension reload and disable", async () => {
  let nextId = 0;
  let queued: { id: string; type: string; payload: object } | undefined;
  let receive: ((response: BridgeResponse) => void) | undefined;
  const opened = await openTaskpane(env, {
    bridge: async (url, route) => {
      if (url.pathname === "/client/poll" && queued) {
        const command = queued;
        queued = undefined;
        await route.fulfill({ contentType: "application/json", body: JSON.stringify(command) });
        return true;
      }
      if (url.pathname === "/client/result") {
        const raw: unknown = JSON.parse(route.request().postData() ?? "{}");
        assert.ok(Value.Check(responseSchema, raw), "valid bridge response");
        receive?.(raw);
        await route.fulfill({ contentType: "application/json", body: JSON.stringify({ ok: true }) });
        return true;
      }
      return false;
    },
  });
  const command = async (type: string, payload: object = {}): Promise<BridgeResponse> => {
    const id = String(++nextId);
    return new Promise<BridgeResponse>((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error(`Bridge ${type} timed out`)), 10_000);
      receive = (response) => {
        if (response.commandId !== id) return;
        clearTimeout(timer);
        resolve(response);
      };
      queued = { id, type, payload };
    });
  };
  try {
    const installed = await command("extensionInstallCode", {
      name: "Simple counter",
      code: `export function activate(api) {
        api.registerTool('durable_counter_simple', {
          description: 'Increment a test counter exactly once when user asks.',
          parameters: { type: 'object', properties: {}, additionalProperties: false },
          async execute(params) {
            const count = (await api.storage.get('count') || 0) + 1;
            await api.storage.set('count', count);
            return { content: [{ type: 'text', text: 'durable counter ' + count }], details: { count } };
          }
        });
      }`,
    });
    assert.equal(installed.ok, true);
    assert.ok(Value.Check(installedSchema, installed.result));
    const { extensionId } = installed.result;
    const listed = await command("extensionList");
    assert.equal(listed.ok, true);
    assert.ok(Value.Check(extensionsSchema, listed.result));
    const extension = listed.result.extensions.find((entry) => entry.name === "Simple counter");
    assert.ok(extension);
    assert.equal(extension.loaded, true);
    assert.equal(extension.runtimeMode, "sandbox-iframe");
    assert.equal(extension.lastError, null);

    for (const count of [1, 2, 3]) {
      if (count === 2) {
        await opened.page.reload();
        await opened.page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
      }
      if (count === 3) assert.equal((await command("extensionReload", { extensionId })).ok, true);
      const executed = await command("executeTool", { name: "durable_counter_simple", args: {} });
      assert.equal(executed.ok, true, executed.error?.message);
      assert.ok(Value.Check(counterResultSchema, executed.result));
      assert.equal(executed.result.details.count, count);
      assert.deepEqual(executed.result.content, [{ type: "text", text: `durable counter ${count}` }]);
      const stored = await command("extensionStorageGet", { extensionId, key: "count" });
      assert.equal(stored.ok, true);
      assert.ok(Value.Check(storageResultSchema, stored.result));
      assert.equal(stored.result.value, count, "read back the independently persisted counter");
    }

    assert.equal((await command("extensionSetEnabled", { extensionId, enabled: false })).ok, true);
    const unavailable = await command("executeTool", { name: "durable_counter_simple", args: {} });
    assert.equal(unavailable.ok, false);
    assert.equal(unavailable.error?.message, "Tool not available: durable_counter_simple");
    const unchanged = await command("extensionStorageGet", { extensionId, key: "count" });
    assert.ok(Value.Check(storageResultSchema, unchanged.result));
    assert.equal(unchanged.result.value, 3, "disabled tools cannot execute stale handlers");
  } finally { await opened.finish(); }
});
