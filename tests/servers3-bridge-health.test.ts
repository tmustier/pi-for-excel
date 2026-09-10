import assert from "node:assert/strict";
import { createServer, type Server } from "node:http";
import { test } from "node:test";

import { createExperimentalCommands } from "../src/commands/builtins/experimental.ts";

interface HealthResponse {
  status: number;
  body: string;
}

async function listen(server: Server): Promise<string> {
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  if (address === null || typeof address === "string") throw new Error("Expected TCP server address.");
  return `http://127.0.0.1:${address.port}`;
}

async function close(server: Server): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    server.close((error) => error ? reject(error) : resolve());
  });
}

function experimentalCommand(dependencies: Parameters<typeof createExperimentalCommands>[0]) {
  const command = createExperimentalCommands(dependencies).find((entry) => entry.name === "experimental");
  if (!command) throw new Error("Experimental command is not registered.");
  return command;
}

void test("/experimental bridge diagnostics describe each HTTP health state", async () => {
  let response: HealthResponse = { status: 200, body: "{}" };
  const server = createServer((_request, outgoing) => {
    outgoing.writeHead(response.status, { "content-type": "application/json" });
    outgoing.end(response.body);
  });
  const bridgeUrl = await listen(server);

  try {
    const rows: ReadonlyArray<{ response: HealthResponse; expected: RegExp }> = [
      {
        response: { status: 200, body: JSON.stringify({ ok: true, mode: "tmux", backend: "tmux", sessions: 2 }) },
        expected: /gate: pass[\s\S]*health: reachable \(HTTP 200, mode=tmux, backend=tmux, sessions=2\)/u,
      },
      {
        response: { status: 200, body: JSON.stringify({ ok: true, mode: "stub", backend: "stub", sessions: 0 }) },
        expected: /health: reachable \(HTTP 200, mode=stub, backend=stub, sessions=0\)/u,
      },
      {
        response: { status: 503, body: JSON.stringify({ error: "tmux unavailable" }) },
        expected: /gate: blocked \(bridge_unreachable\)[\s\S]*health: unreachable \(HTTP 503; tmux unavailable\)/u,
      },
      {
        response: { status: 200, body: "not json" },
        expected: /gate: pass[\s\S]*health: reachable \(HTTP 200\)/u,
      },
    ];

    for (const row of rows) {
      response = row.response;
      const toasts: string[] = [];
      const command = experimentalCommand({
        showExperimentalDialog: () => {},
        showToast: (message) => toasts.push(message),
        getTmuxBridgeUrl: () => Promise.resolve(bridgeUrl),
        getTmuxBridgeToken: () => Promise.resolve(undefined),
      });

      await command.execute("tmux-status");
      assert.match(toasts[0] ?? "", row.expected);
    }
  } finally {
    await close(server);
  }
});

void test("bridge-status alias wires diagnostics to the bridge /health endpoint", async () => {
  let requestedPath = "";
  const server = createServer((request, response) => {
    requestedPath = request.url ?? "";
    response.writeHead(200, { "content-type": "application/json" });
    response.end(JSON.stringify({ ok: true, mode: "tmux" }));
  });
  const bridgeUrl = await listen(server);

  try {
    const toasts: string[] = [];
    const command = experimentalCommand({
      showExperimentalDialog: () => {},
      showToast: (message) => toasts.push(message),
      getTmuxBridgeUrl: () => Promise.resolve(bridgeUrl),
      getTmuxBridgeToken: () => Promise.resolve(undefined),
    });

    await command.execute("bridge-status");

    assert.deepEqual({ requestedPath, diagnostic: toasts[0]?.split("\n")[0] }, {
      requestedPath: "/health",
      diagnostic: "Tmux bridge status:",
    });
  } finally {
    await close(server);
  }
});
