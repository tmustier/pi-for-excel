import assert from "node:assert/strict";
import http from "node:http";
import os from "node:os";
import test from "node:test";

import WebSocket, { WebSocketServer } from "ws";

import {
  bridgeCodexWebSocketToSse,
  CODEX_WEBSOCKET_BRIDGE_HEADER,
} from "../scripts/codex-websocket-bridge.mjs";

async function listen(server) {
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  const address = server.address();
  if (!address || typeof address === "string") {
    throw new Error("Expected TCP server address");
  }
  return address.port;
}

async function close(server) {
  await new Promise((resolve) => server.close(resolve));
}

function localWebSocketConstructor(port) {
  return class LocalWebSocket extends WebSocket {
    constructor(_url, options) {
      super(`ws://127.0.0.1:${port}/backend-api/codex/responses`, options);
    }
  };
}

const CODEX_TARGET_URL = new URL("https://chatgpt.com/backend-api/codex/responses");

test("Codex WebSocket bridge converts upstream JSON frames to SSE", async (t) => {
  let receivedBody = "";
  let receivedHeaders = {};
  const upstreamServer = http.createServer();
  const upstreamWebSocket = new WebSocketServer({ server: upstreamServer });
  upstreamWebSocket.on("connection", (socket, request) => {
    receivedHeaders = request.headers;
    socket.on("message", (data) => {
      receivedBody = data.toString("utf8");
      socket.send(JSON.stringify({ type: "response.created", response: { id: "resp_test" } }));
      socket.send(JSON.stringify({ type: "response.output_text.delta", delta: "LUNA_OK" }, null, 2));
      socket.send(JSON.stringify({ type: "response.completed", response: { id: "resp_test" } }));
    });
  });
  const upstreamPort = await listen(upstreamServer);

  const bridgeServer = http.createServer((req, res) => {
    void bridgeCodexWebSocketToSse({
      req,
      res,
      targetUrl: CODEX_TARGET_URL,
      outboundHeaders: new Headers({
        Authorization: "Bearer test-token",
        "chatgpt-account-id": "account-test",
        Cookie: "must-not-forward=1",
        "Content-Type": "application/json",
        "OpenAI-Beta": "responses=experimental",
        Origin: "https://localhost:3141",
        Referer: "https://localhost:3141/src/taskpane.html",
        "Sec-WebSocket-Protocol": "must-not-forward",
      }),
      WebSocketConstructor: localWebSocketConstructor(upstreamPort),
    });
  });
  const bridgePort = await listen(bridgeServer);

  t.after(async () => {
    upstreamWebSocket.close();
    await close(bridgeServer);
    await close(upstreamServer);
  });

  const requestPayload = { model: "gpt-5.6-luna", stream: true, input: [] };
  const response = await fetch(`http://127.0.0.1:${bridgePort}/`, {
    method: "POST",
    body: JSON.stringify(requestPayload),
  });

  assert.equal(response.status, 200);
  assert.equal(response.headers.get(CODEX_WEBSOCKET_BRIDGE_HEADER), "1");
  assert.match(response.headers.get("content-type") ?? "", /text\/event-stream/);

  const body = await response.text();
  assert.match(body, /data: \{"type":"response\.created"/);
  assert.match(body, /data: \{"type":"response\.output_text\.delta","delta":"LUNA_OK"\}/);
  assert.match(body, /data: \{"type":"response\.completed"/);
  assert.deepEqual(JSON.parse(receivedBody), { ...requestPayload, type: "response.create" });
  assert.equal(receivedHeaders.authorization, "Bearer test-token");
  assert.equal(receivedHeaders["chatgpt-account-id"], "account-test");
  assert.equal(receivedHeaders["openai-beta"], "responses_websockets=2026-02-06");
  assert.equal(receivedHeaders.originator, "pi");
  assert.equal(receivedHeaders["user-agent"], `pi (${os.platform()} ${os.release()}; ${os.arch()})`);
  assert.equal(receivedHeaders["content-type"], undefined);
  assert.equal(receivedHeaders.cookie, undefined);
  assert.equal(receivedHeaders.origin, undefined);
  assert.equal(receivedHeaders.referer, undefined);
  assert.equal(receivedHeaders["sec-websocket-protocol"], undefined);
});

test("Codex WebSocket bridge emits an SSE error when upstream closes before a terminal event", async (t) => {
  const upstreamServer = http.createServer();
  const upstreamWebSocket = new WebSocketServer({ server: upstreamServer });
  upstreamWebSocket.on("connection", (socket) => {
    socket.on("message", () => {
      socket.send(JSON.stringify({ type: "response.created", response: { id: "resp_partial" } }));
      socket.send(JSON.stringify({ type: "response.output_text.delta", delta: "PARTIAL" }));
      socket.close(1000, "premature");
    });
  });
  const upstreamPort = await listen(upstreamServer);

  const bridgeServer = http.createServer((req, res) => {
    void bridgeCodexWebSocketToSse({
      req,
      res,
      targetUrl: CODEX_TARGET_URL,
      outboundHeaders: new Headers(),
      WebSocketConstructor: localWebSocketConstructor(upstreamPort),
    });
  });
  const bridgePort = await listen(bridgeServer);

  t.after(async () => {
    upstreamWebSocket.close();
    await close(bridgeServer);
    await close(upstreamServer);
  });

  const response = await fetch(`http://127.0.0.1:${bridgePort}/`, {
    method: "POST",
    body: JSON.stringify({ model: "gpt-5.6-luna", stream: true, input: [] }),
  });
  const events = (await response.text())
    .split("\n")
    .filter((line) => line.startsWith("data: "))
    .map((line) => JSON.parse(line.slice(6)));

  assert.equal(response.status, 200);
  assert.deepEqual(events.map((event) => event.type), [
    "response.created",
    "response.output_text.delta",
    "error",
  ]);
  assert.equal(events[2].error.type, "proxy_stream_error");
  assert.match(events[2].error.message, /closed before a terminal response event/);
});

test("Codex WebSocket bridge rejects a non-ChatGPT target inside the helper", async (t) => {
  const bridgeServer = http.createServer((req, res) => {
    void bridgeCodexWebSocketToSse({
      req,
      res,
      targetUrl: new URL("https://example.com/backend-api/codex/responses"),
      outboundHeaders: new Headers(),
    });
  });
  const bridgePort = await listen(bridgeServer);
  t.after(() => close(bridgeServer));

  const response = await fetch(`http://127.0.0.1:${bridgePort}/`, {
    method: "POST",
    body: "{}",
  });
  assert.equal(response.status, 400);
  assert.match(await response.text(), /target is not allowed/);
});

test("Codex WebSocket bridge rejects non-POST requests", async (t) => {
  const bridgeServer = http.createServer((req, res) => {
    void bridgeCodexWebSocketToSse({
      req,
      res,
      targetUrl: new URL("http://127.0.0.1:9/backend-api/codex/responses"),
      outboundHeaders: new Headers(),
    });
  });
  const bridgePort = await listen(bridgeServer);
  t.after(() => close(bridgeServer));

  const response = await fetch(`http://127.0.0.1:${bridgePort}/`);
  assert.equal(response.status, 405);
  assert.match(await response.text(), /requires POST/);
});
