import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import { Value } from "typebox/value";
import { openTaskpane, startTaskpaneServer } from "./harness.ts";

void test("the dev bridge configures MCP only in the active chat and rejects invalid settings", async () => {
  const env = await startTaskpaneServer();
  let nextCommand: { id: string; type: string; payload: Record<string, string | boolean> } | null = null;
  let receive = (_body: string): void => {};
  const opened = await openTaskpane(env, {
    clientId: "mcp-integration-setup",
    bridge: async (url, route) => {
      if (url.pathname === "/client/poll") {
        const command = nextCommand ?? { type: "noop" };
        nextCommand = null;
        if (command.type === "noop") await new Promise((resolve) => setTimeout(resolve, 50));
        await route.fulfill({ contentType: "application/json", body: JSON.stringify(command) });
        return true;
      }
      if (url.pathname === "/client/result") {
        receive(route.request().postData() ?? "");
        await route.fulfill({ contentType: "application/json", body: '{"ok":true}' });
        return true;
      }
      return false;
    },
  });
  async function send(payload: Record<string, string | boolean>): Promise<string> {
    const result = new Promise<string>((resolve) => { receive = resolve; });
    nextCommand = { id: "configure", type: "configureIntegration", payload };
    let timer: ReturnType<typeof setTimeout> | undefined;
    // agent-default (not a user rule): 2026-10-17 — fail a stalled test bridge after 10 seconds.
    const timeout = new Promise<never>((_resolve, reject) => {
      timer = setTimeout(() => reject(new Error("Integration bridge command timed out")), 10_000);
    });
    try { return await Promise.race([result, timeout]); }
    finally { clearTimeout(timer); }
  }
  const hasMcp = 'document.querySelector("pi-sidebar").agent.state.tools.some(tool => tool.name === "mcp")';
  try {
    const enabled: unknown = JSON.parse(await send({ integrationId: "mcp_tools", enabled: true }));
    assert.ok(Value.Check(Type.Object({
      ok: Type.Literal(true), result: Type.Object({ configured: Type.Literal(true), sessionId: Type.String() }),
    }), enabled));
    await opened.page.waitForFunction(hasMcp);
    assert.deepEqual(await opened.page.evaluate(`(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts");
      return getAppStorage().settings.get("integrations.session.v1." + ${JSON.stringify(enabled.result.sessionId)});
    })()`), ["mcp_tools"]);
    assert.match(await send({ integrationId: "not-an-integration", enabled: true }), /"ok":false/);
    assert.match(await send({ integrationId: "mcp_tools", enabled: "true" }), /"ok":false/);
    assert.match(await send({ integrationId: "mcp_tools", enabled: true, scope: "workbook" }), /"ok":false/);
    await opened.page.getByRole("button", { name: "New tab" }).click();
    await opened.page.waitForFunction(`!(${hasMcp})`);
    await opened.page.getByRole("button", { name: "Open tab Chat 1", exact: true }).click();
    await opened.page.waitForFunction(hasMcp);
    assert.match(await send({ integrationId: "mcp_tools", enabled: false }), /"configured":true/);
    await opened.page.waitForFunction(`!(${hasMcp})`);
  } finally {
    await opened.finish();
    await env.close();
  }
});
