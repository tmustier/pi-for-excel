import assert from "node:assert/strict";
import { test } from "node:test";

import type {
  Api,
  Model,
} from "@earendil-works/pi-ai";
import { createModels } from "@earendil-works/pi-ai";
import { builtinModels } from "@earendil-works/pi-ai/providers/all";

import { BROWSER_OAUTH_PROVIDERS, mapToApiProvider } from "../src/auth/provider-map.ts";
import { rewriteDevProxyUrl } from "../src/auth/dev-rewrites.ts";
import { installBedrockProviderStub } from "../src/compat/bedrock-provider-stub.ts";
import { installProcessEnvShim } from "../src/compat/process-env-shim.ts";
import { orderModelsForSelector } from "../src/models/featured-models.ts";
import {
  compareModels,
  parseMajorMinor,
} from "../src/models/model-ordering.ts";
import { getThinkingLevelsForModel } from "../src/models/thinking-levels.ts";
import { pickDefaultModel as pickDefaultModelFromRuntime } from "../src/taskpane/default-model.ts";

type OpenAiProvider = "openai" | "openai-codex";

const modelsRuntime = builtinModels();

function getModels(provider: string): Model<Api>[] {
  return [...modelsRuntime.getModels(provider)];
}

function getModel(provider: string, modelId: string): Model<Api> {
  const model = modelsRuntime.getModel(provider, modelId);
  if (!model) throw new Error(`Missing test model: ${provider}/${modelId}`);
  return model;
}

function pickDefaultModel(availableProviders: string[]): Model<Api> {
  return pickDefaultModelFromRuntime(modelsRuntime, availableProviders);
}

const OPENAI_PROVIDERS: OpenAiProvider[] = ["openai", "openai-codex"];

const GPT_56_IDS = ["gpt-5.6-sol", "gpt-5.6-terra", "gpt-5.6-luna"];

void test("parseMajorMinor handles registry version formats", () => {
  const cases: Array<[string, number]> = [
    ["claude-opus-4-6", 46],
    ["anthropic/claude-fable-5", 50],
    ["anthropic.claude-opus-4-1-20250805-v1:0", 41],
    ["claude-opus-4-20250514", 40],
    ["gpt-5.6-sol", 56],
    ["gemini-2.5-pro", 25],
    ["gpt-4o-2024-11-20", 40],
    ["gemini-2.5-pro-preview-06-05", 25],
    ["gpt-5.12", 512],
    ["gemma-4-31b-it", 40],
    ["zai.glm-4.7", 47],
    ["deepseek.v3.2", 32],
    ["unknown-model-2024-11-20", 0],
  ];

  for (const [id, expected] of cases) assert.equal(parseMajorMinor(id), expected, id);
});

void test("model selector orders all GPT-5.6 tiers as Sol, Terra, Luna", () => {
  const items = GPT_56_IDS
    .slice()
    .reverse()
    .map((id) => ({ provider: "openai-codex", id, model: getModel("openai-codex", id) }));

  const ordered = orderModelsForSelector(items, null);
  assert.deepEqual(ordered.map((item) => item.id), GPT_56_IDS);
});

void test("GPT-6.1 Sol thinking controls omit unsupported levels", () => {
  for (const provider of OPENAI_PROVIDERS) {
    assert.deepEqual(
      getThinkingLevelsForModel(getModel(provider, "gpt-6.1-sol")),
      provider === "openai-codex"
        ? ["minimal", "low", "medium", "high", "xhigh", "max"]
        : ["low", "medium", "high", "xhigh", "max"],
    );
  }
});

void test("model selector orders GPT-6 tiers as Sol, Astra, Luna", () => {
  const ids = ["gpt-6-luna", "gpt-6-astra", "gpt-6-sol"];
  const items = ids.map((id) => ({ provider: "openai-codex", id, model: getModel("openai-codex", id) }));

  const ordered = orderModelsForSelector(items, null);
  assert.deepEqual(ordered.map((item) => item.id), ["gpt-6-sol", "gpt-6-astra", "gpt-6-luna"]);
});

void test("pickDefaultModel prefers the latest Opus for Anthropic-only setups", () => {
  const selected = pickDefaultModel(["anthropic"]);
  assert.equal(selected.provider, "anthropic");
  assert.equal(selected.id, "claude-opus-5-5");
});

void test("current OpenAI providers select GPT-6.1 Sol as the default", () => {
  for (const provider of OPENAI_PROVIDERS) {
    const selected = pickDefaultModel([provider]);
    assert.equal(selected.provider, provider);
    assert.equal(selected.id, "gpt-6.1-sol");
  }
});

void test("OpenAI defaults compare versions before tiers and keep Codex as a fallback", () => {
  const provider = modelsRuntime.getProvider("openai");
  assert.ok(provider);
  const cases: Array<{ ids: string[]; expected: string }> = [
    { ids: ["gpt-6-luna", "gpt-6-astra", "gpt-6-sol"], expected: "gpt-6-sol" },
    { ids: ["gpt-6-sol", "gpt-6.1-sol"], expected: "gpt-6.1-sol" },
    { ids: ["gpt-5.3-codex", "gpt-6-luna"], expected: "gpt-6-luna" },
    { ids: ["gpt-5.3-codex", "gpt-5.2"], expected: "gpt-5.3-codex" },
    { ids: ["gpt-5.3-codex"], expected: "gpt-5.3-codex" },
  ];
  for (const { ids, expected } of cases) {
    const runtime = createModels();
    runtime.setProvider({ ...provider, getModels: () => ids.map((id) => getModel("openai", id)) });
    assert.equal(pickDefaultModelFromRuntime(runtime, ["openai"]).id, expected);
  }
});

void test("Bedrock provider uses the browser-safe unsupported-provider stub", async () => {
  installBedrockProviderStub();

  const selected = await modelsRuntime.completeSimple(
    getModel("amazon-bedrock", "amazon.nova-micro-v1:0"),
    {
      messages: [
        {
          role: "user",
          content: [{ type: "text", text: "hello" }],
          timestamp: Date.now(),
        },
      ],
    },
    { apiKey: "browser-test", maxTokens: 1, maxRetries: 0 },
  );

  assert.equal(selected.stopReason, "error");
  assert.match(selected.errorMessage ?? "", /Amazon Bedrock is not supported/);
});

void test("pickDefaultModel falls back to the preferred hardcoded OpenAI default", () => {
  const selected = pickDefaultModel([]);
  assert.equal(selected.provider, "openai");
  assert.equal(selected.id, "gpt-6.1-sol");
});

void test("pickDefaultModel never picks an unusable provider while a configured provider has models (#553)", () => {
  // Providers with registry models but no dedicated default-model rule used
  // to fall through to an OpenAI API model, which the user has no credentials for.
  for (const provider of ["github-copilot", "mistral", "groq", "xai", "deepseek"]) {
    const models = getModels(provider);
    assert.ok(models.length > 0, `expected registry models for ${provider}`);

    const selected = pickDefaultModel([provider]);
    assert.equal(
      selected.provider,
      provider,
      `expected default model from ${provider}, got ${selected.provider}/${selected.id}`,
    );
  }
});

void test("pickDefaultModel tolerates unknown provider names", () => {
  const selected = pickDefaultModel(["some-custom-gateway"]);
  assert.equal(selected.provider, "openai");
  assert.equal(selected.id, "gpt-6.1-sol");
});

void test("pickDefaultModel prefers GPT-6.1 Sol when OpenAI and Anthropic are both available", () => {
  const selected = pickDefaultModel(["anthropic", "openai"]);
  assert.equal(selected.provider, "openai");
  assert.equal(selected.id, "gpt-6.1-sol");
});

void test("openai compareModels does not mistake dated GPT-4o ids for newer versions", () => {
  const models = [
    { provider: "openai", id: "gpt-4o-2024-11-20" },
    { provider: "openai", id: "gpt-5.5" },
    { provider: "openai", id: "gpt-5.4-pro" },
    { provider: "openai", id: "gpt-5.3-codex" },
  ];

  models.sort(compareModels);

  assert.deepEqual(models.map((m) => m.id), ["gpt-5.5", "gpt-5.4-pro", "gpt-5.3-codex", "gpt-4o-2024-11-20"]);
});

void test("provider-map keeps openai-codex distinct from openai", () => {
  assert.equal(mapToApiProvider("openai-codex"), "openai-codex");
  assert.equal(mapToApiProvider("openai"), "openai");
});

void test("provider-map keeps Google OAuth providers distinct from API-key google", () => {
  assert.equal(mapToApiProvider("gemini-cli"), "google-gemini-cli");
  assert.equal(mapToApiProvider("google-gemini-cli"), "google-gemini-cli");
  assert.equal(mapToApiProvider("antigravity"), "google-antigravity");
  assert.equal(mapToApiProvider("google-antigravity"), "google-antigravity");
});

void test("browser oauth providers include OpenAI + Google OAuth providers", () => {
  assert.equal(BROWSER_OAUTH_PROVIDERS.includes("openai-codex"), true);
  assert.equal(BROWSER_OAUTH_PROVIDERS.includes("google-gemini-cli"), true);
  assert.equal(BROWSER_OAUTH_PROVIDERS.includes("google-antigravity"), true);
});

void test("process-env shim adds process.env for browser-like runtimes", () => {
  const runtime: { process?: unknown } = {};
  installProcessEnvShim(runtime);

  assert.ok(runtime.process && typeof runtime.process === "object" && !Array.isArray(runtime.process));

  assert.ok("env" in runtime.process);

  const envValue = runtime.process.env;
  assert.ok(envValue && typeof envValue === "object" && !Array.isArray(envValue));
});

void test("dev rewrite routes OAuth hosts to dedicated proxies", () => {
  assert.equal(
    rewriteDevProxyUrl("https://platform.claude.com/v1/oauth/token"),
    "/oauth-proxy/anthropic-platform/v1/oauth/token",
  );

  assert.equal(
    rewriteDevProxyUrl("https://console.anthropic.com/v1/oauth/token"),
    "/oauth-proxy/anthropic/v1/oauth/token",
  );
});

void test("dev rewrite routes cloudcode hosts to dedicated proxies", () => {
  assert.equal(
    rewriteDevProxyUrl("https://cloudcode-pa.googleapis.com/v1internal:streamGenerateContent?alt=sse"),
    "/api-proxy/google-cloudcode/v1internal:streamGenerateContent?alt=sse",
  );

  assert.equal(
    rewriteDevProxyUrl("https://daily-cloudcode-pa.sandbox.googleapis.com/v1internal:streamGenerateContent?alt=sse"),
    "/api-proxy/google-cloudcode-sandbox/v1internal:streamGenerateContent?alt=sse",
  );

  assert.equal(
    rewriteDevProxyUrl("https://generativelanguage.googleapis.com/v1beta/models"),
    "/api-proxy/google/v1beta/models",
  );

  assert.equal(rewriteDevProxyUrl("https://example.com/test"), null);
});
