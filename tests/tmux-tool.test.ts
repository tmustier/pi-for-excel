import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@mariozechner/pi-agent-core";

import {
  computeTmuxFetchTimeoutMs,
  createTmuxTool,
  type TmuxBridgeConfig,
  type TmuxBridgeRequest,
  type TmuxBridgeResponse,
  type TmuxToolDetails,
} from "../src/tools/tmux.ts";

function firstText(result: AgentToolResult<TmuxToolDetails>): string {
  const first = result.content[0];
  if (!first || first.type !== "text") {
    throw new Error("Expected first content block to be text");
  }
  return first.text;
}

void test("tmux fetch timeout includes wait_ms and send_and_capture timeout", () => {
  const captureTimeout = computeTmuxFetchTimeoutMs({
    action: "capture_pane",
    session: "demo",
    wait_ms: 30_000,
  });
  assert.equal(captureTimeout, 35_000);

  const sendAndCaptureTimeout = computeTmuxFetchTimeoutMs({
    action: "send_and_capture",
    session: "demo",
    text: "pi --help",
    wait_ms: 30_000,
    timeout_ms: 120_000,
  });
  assert.equal(sendAndCaptureTimeout, 155_000);
});

void test("tmux fetch timeout keeps sane defaults and hard cap", () => {
  const defaultTimeout = computeTmuxFetchTimeoutMs({
    action: "send_and_capture",
    session: "demo",
    text: "echo ready",
  });
  assert.equal(defaultTimeout, 15_000);

  const cappedTimeout = computeTmuxFetchTimeoutMs({
    action: "send_and_capture",
    session: "demo",
    text: "echo done",
    wait_ms: 240_000,
    timeout_ms: 120_000,
  });
  assert.equal(cappedTimeout, 245_000);
});

void test("tmux tool returns guidance when bridge URL is not configured", async () => {
  const tool = createTmuxTool({
    getBridgeConfig: () => Promise.resolve(null),
  });

  const result = await tool.execute("tc-missing", { action: "list_sessions" });

  assert.match(firstText(result), /tmux-bridge-url/u);
  assert.match(firstText(result), /https:\/\/localhost:3341/u);
  assert.equal(result.details?.kind, "tmux_bridge");
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "missing_bridge_url");
});

void test("tmux tool validates required session/action params", async () => {
  const tool = createTmuxTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3341" }),
    callBridge: () => Promise.resolve({ ok: true, action: "send_keys" }),
  });

  const result = await tool.execute("tc-validate", {
    action: "send_keys",
    text: "ls",
  });

  assert.match(firstText(result), /session is required/u);
  assert.equal(result.details?.ok, false);
});

void test("tmux tool sends v1 bridge contract payload for send_and_capture", async () => {
  let capturedRequest: TmuxBridgeRequest | null = null;
  let capturedConfig: TmuxBridgeConfig | null = null;

  const tool = createTmuxTool({
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3341",
      token: "secret-token",
    }),
    callBridge: (
      request: TmuxBridgeRequest,
      config: TmuxBridgeConfig,
    ): Promise<TmuxBridgeResponse> => {
      capturedRequest = request;
      capturedConfig = config;

      return Promise.resolve({
        ok: true,
        action: "send_and_capture",
        session: "dev-session",
        output: "ready",
      });
    },
  });

  const result = await tool.execute("tc-contract", {
    action: "send_and_capture",
    session: "dev-session",
    text: "npm test",
    enter: true,
    lines: 80,
    wait_for: "done",
    timeout_ms: 5000,
    wait_ms: 15000,
    join_wrapped: true,
  });

  assert.ok(capturedRequest);
  assert.equal(capturedRequest?.action, "send_and_capture");
  assert.equal(capturedRequest?.session, "dev-session");
  assert.equal(capturedRequest?.text, "npm test");
  assert.equal(capturedRequest?.enter, true);
  assert.equal(capturedRequest?.lines, 80);
  assert.equal(capturedRequest?.wait_for, "done");
  assert.equal(capturedRequest?.timeout_ms, 5000);
  assert.equal(capturedRequest?.wait_ms, 15000);
  assert.equal(capturedRequest?.join_wrapped, true);

  assert.ok(capturedConfig);
  assert.equal(capturedConfig?.url, "https://localhost:3341");
  assert.equal(capturedConfig?.token, "secret-token");

  const text = firstText(result);
  assert.match(text, /Sent keys and captured tmux pane/u);
  assert.match(text, /```/u);
  assert.equal(result.details?.ok, true);
  assert.equal(result.details?.action, "send_and_capture");
  assert.equal(result.details?.session, "dev-session");
});

void test("tmux list_sessions response is rendered as a bullet list", async () => {
  const tool = createTmuxTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3341" }),
    callBridge: () => Promise.resolve({
      ok: true,
      action: "list_sessions",
      sessions: ["dev", "ops"],
    }),
  });

  const result = await tool.execute("tc-list", { action: "list_sessions" });

  const text = firstText(result);
  assert.match(text, /Tmux sessions:/u);
  assert.match(text, /- dev/u);
  assert.match(text, /- ops/u);
  assert.equal(result.details?.sessionsCount, 2);
});

void test("tmux bridge errors are surfaced to the user", async () => {
  const tool = createTmuxTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3341" }),
    callBridge: () => Promise.reject(new Error("bridge unavailable")),
  });

  const result = await tool.execute("tc-error", {
    action: "capture_pane",
    session: "dev",
  });

  assert.equal(firstText(result), "Error: bridge unavailable");
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.action, "capture_pane");
  assert.equal(result.details?.error, "bridge unavailable");
});

void test("tmux tool handles explicit bridge-level rejection payloads", async () => {
  const tool = createTmuxTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3341" }),
    callBridge: () => Promise.resolve({
      ok: false,
      action: "kill_session",
      error: "session not found",
    }),
  });

  const result = await tool.execute("tc-reject", {
    action: "kill_session",
    session: "missing",
  });

  assert.equal(firstText(result), "Error: session not found");
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "session not found");
});
