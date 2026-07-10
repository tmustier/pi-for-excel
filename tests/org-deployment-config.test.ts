/**
 * Tests for org/central-deployment build-time configuration:
 * - VITE_PI_DEFAULT_PROXY_URL → resolveDefaultProxyUrl()
 * - VITE_PI_ALLOWED_PROVIDERS → provider allowlist filtering
 */

import test from "node:test";
import assert from "node:assert/strict";

import {
  CODEX_WEBSOCKET_BRIDGE_HEADER,
  DEFAULT_LOCAL_PROXY_URL,
  PROXY_HEALTH_HEADER,
  WPS_DEV_HOST_GATEWAY_PROXY_URL,
  probeCodexWebSocketBridge,
  resolveDefaultProxyUrl,
  resolveRuntimeDefaultProxyUrl,
} from "../src/auth/proxy-validation.ts";
import {
  requiresCodexWebSocketBridge,
  resolveCodexWebSocketBridgeSessionId,
} from "../src/auth/stream-proxy.ts";
import { filterProvidersByAllowlist, resolveAllowedProviderIds } from "../src/ui/provider-allowlist.ts";

void test("resolveDefaultProxyUrl falls back to local default when unset", () => {
  assert.equal(resolveDefaultProxyUrl(undefined), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl(""), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl("   "), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl(42), DEFAULT_LOCAL_PROXY_URL);
});

void test("resolveDefaultProxyUrl accepts https URLs and strips trailing slashes", () => {
  assert.equal(
    resolveDefaultProxyUrl("https://pi-proxy.example.com:3003"),
    "https://pi-proxy.example.com:3003",
  );
  assert.equal(
    resolveDefaultProxyUrl("https://pi-proxy.example.com:3003/"),
    "https://pi-proxy.example.com:3003",
  );
});

void test("resolveDefaultProxyUrl refuses http (mixed content) and garbage", () => {
  assert.equal(resolveDefaultProxyUrl("http://pi-proxy.example.com:3003"), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl("pi-proxy.example.com"), DEFAULT_LOCAL_PROXY_URL);
  assert.equal(resolveDefaultProxyUrl("https://"), DEFAULT_LOCAL_PROXY_URL);
});

void test("resolveRuntimeDefaultProxyUrl uses host-gateway proxy for WPS HTTP harness", () => {
  assert.equal(
    resolveRuntimeDefaultProxyUrl({ hostKind: "wps", location: { protocol: "http:", hostname: "10.0.2.2" } }),
    WPS_DEV_HOST_GATEWAY_PROXY_URL,
  );
  assert.equal(
    resolveRuntimeDefaultProxyUrl({ hostKind: "office", location: { protocol: "http:", hostname: "10.0.2.2" } }),
    DEFAULT_LOCAL_PROXY_URL,
  );
  assert.equal(
    resolveRuntimeDefaultProxyUrl({ hostKind: "wps", location: { protocol: "https:", hostname: "10.0.2.2" } }),
    DEFAULT_LOCAL_PROXY_URL,
  );
});

void test("Codex WebSocket bridge capability probe requires the advertised health header", async () => {
  const previousFetch = globalThis.fetch;
  const seenUrls: string[] = [];

  try {
    Reflect.set(globalThis, "fetch", (input: RequestInfo | URL) => {
      const url = typeof input === "string"
        ? input
        : input instanceof URL
          ? input.toString()
          : input.url;
      seenUrls.push(url);
      return Promise.resolve(new Response("ok", {
        status: 200,
        headers: {
          [PROXY_HEALTH_HEADER]: "1",
          [CODEX_WEBSOCKET_BRIDGE_HEADER]: "1",
        },
      }));
    });

    assert.equal(await probeCodexWebSocketBridge("https://localhost:3003/"), true);
    assert.deepEqual(seenUrls, ["https://localhost:3003/healthz"]);

    Reflect.set(globalThis, "fetch", () => Promise.resolve(new Response("ok", { status: 200 })));
    assert.equal(await probeCodexWebSocketBridge("https://localhost:3003"), false);
  } finally {
    Reflect.set(globalThis, "fetch", previousFetch);
  }
});

void test("only ChatGPT GPT-5.6 Luna requires the Codex WebSocket bridge", () => {
  assert.equal(requiresCodexWebSocketBridge({ provider: "openai-codex", id: "gpt-5.6-luna" }), true);
  assert.equal(requiresCodexWebSocketBridge({ provider: "openai-codex", id: "gpt-5.6-sol" }), false);
  assert.equal(requiresCodexWebSocketBridge({ provider: "openai", id: "gpt-5.6-luna" }), false);
});

void test("Codex WebSocket bridge preserves UUIDv7 and stably maps legacy session ids", () => {
  const nativeSessionId = "019f4c1c-03ae-7d15-8e28-035d6a58c787";
  assert.equal(resolveCodexWebSocketBridgeSessionId(nativeSessionId), nativeSessionId);

  const legacySessionId = "b01d800c-e36c-4737-b987-c5ebb16d4106";
  const mapped = resolveCodexWebSocketBridgeSessionId(legacySessionId);
  assert.match(mapped, /^[0-9a-f]{8}-[0-9a-f]{4}-7[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/u);
  assert.equal(resolveCodexWebSocketBridgeSessionId(legacySessionId), mapped);
  assert.notEqual(
    resolveCodexWebSocketBridgeSessionId("6a4bafc5-79fc-4c27-a9e3-86bf0c35917f"),
    mapped,
  );
  assert.notEqual(
    resolveCodexWebSocketBridgeSessionId(),
    resolveCodexWebSocketBridgeSessionId(),
  );
});

const PROVIDERS = [
  { id: "anthropic", label: "Anthropic" },
  { id: "openai", label: "OpenAI (API)" },
  { id: "deepseek", label: "DeepSeek" },
];

void test("resolveAllowedProviderIds returns null when unset", () => {
  assert.equal(resolveAllowedProviderIds(undefined), null);
  assert.equal(resolveAllowedProviderIds(""), null);
  assert.equal(resolveAllowedProviderIds(" , "), null);
  assert.equal(resolveAllowedProviderIds(7), null);
});

void test("resolveAllowedProviderIds parses and lowercases ids", () => {
  const ids = resolveAllowedProviderIds(" OpenAI, deepseek ,");
  assert.notEqual(ids, null);
  assert.deepEqual([...(ids as Set<string>)].sort(), ["deepseek", "openai"]);
});

void test("filterProvidersByAllowlist passes through with no restriction", () => {
  assert.deepEqual(filterProvidersByAllowlist(PROVIDERS, null), PROVIDERS);
});

void test("filterProvidersByAllowlist keeps only allowlisted providers in order", () => {
  const allowed = resolveAllowedProviderIds("deepseek,openai");
  const filtered = filterProvidersByAllowlist(PROVIDERS, allowed);
  assert.deepEqual(filtered.map((p) => p.id), ["openai", "deepseek"]);
});

void test("filterProvidersByAllowlist fails open on fully mismatched allowlist", () => {
  const allowed = resolveAllowedProviderIds("no-such-provider");
  const filtered = filterProvidersByAllowlist(PROVIDERS, allowed);
  assert.deepEqual(filtered.map((p) => p.id), ["anthropic", "openai", "deepseek"]);
});
