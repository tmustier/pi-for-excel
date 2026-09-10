// Static contract: the Vite aliases checked below are browser bundle dependency restrictions;
// model-ordering behavior cannot prove Node-only provider modules stay excluded from the build.
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import { test } from "node:test";

import type {
  Api,
  Context,
  Model,
  SimpleStreamOptions,
} from "@earendil-works/pi-ai";
import { builtinModels } from "@earendil-works/pi-ai/providers/all";

import { BROWSER_OAUTH_PROVIDERS, mapToApiProvider } from "../src/auth/provider-map.ts";
import { rewriteDevProxyUrl } from "../src/auth/dev-rewrites.ts";
import { installBedrockProviderStub } from "../src/compat/bedrock-provider-stub.ts";
import { installProcessEnvShim } from "../src/compat/process-env-shim.ts";
import { orderModelsForSelector } from "../src/models/featured-models.ts";
import {
  compareModels,
  compareOpenAiModelIds,
  isOpenAiCodexModelId,
  isOpenAiGeneralGptModelId,
  parseMajorMinor,
  shouldPreferOpenAiGeneralModel,
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

function completeSimple(
  model: Model<Api>,
  context: Context,
  options?: SimpleStreamOptions,
) {
  return modelsRuntime.completeSimple(model, context, options);
}

const OPENAI_PROVIDERS: OpenAiProvider[] = ["openai", "openai-codex"];

const GPT_56_VARIANTS = [
  {
    id: "gpt-5.6-sol",
    name: "GPT-5.6 Sol",
    cost: {
      input: 4,
      output: 20,
      cacheRead: 0.4,
      cacheWrite: 5,
      tiers: [{ inputTokensAbove: 272_000, input: 8, output: 30, cacheRead: 0.8, cacheWrite: 10 }],
    },
    codexCost: {
      input: 5,
      output: 30,
      cacheRead: 0.5,
      cacheWrite: 6.25,
      tiers: [{ inputTokensAbove: 272_000, input: 10, output: 45, cacheRead: 1, cacheWrite: 12.5 }],
    },
  },
  {
    id: "gpt-5.6-terra",
    name: "GPT-5.6 Terra",
    cost: {
      input: 2,
      output: 12,
      cacheRead: 0.2,
      cacheWrite: 2.5,
      tiers: [{ inputTokensAbove: 272_000, input: 4, output: 18, cacheRead: 0.4, cacheWrite: 5 }],
    },
  },
  {
    id: "gpt-5.6-luna",
    name: "GPT-5.6 Luna",
    cost: {
      input: 0.2,
      output: 1.2,
      cacheRead: 0.02,
      cacheWrite: 0.25,
      tiers: [{ inputTokensAbove: 272_000, input: 0.4, output: 1.8, cacheRead: 0.04, cacheWrite: 0.5 }],
    },
  },
];

const GPT_56_IDS = GPT_56_VARIANTS.map((variant) => variant.id);
const GPT_56_THINKING_LEVELS: Record<OpenAiProvider, string[]> = {
  openai: ["off", "low", "medium", "high", "xhigh", "max"],
  "openai-codex": ["off", "minimal", "low", "medium", "high", "xhigh", "max"],
};

function pickExpectedOpenAiDefault(provider: OpenAiProvider): Model<Api> | null {
  const models = getModels(provider);
  const bestGeneral = models
    .filter((model) => isOpenAiGeneralGptModelId(model.id))
    .sort((a, b) => compareOpenAiModelIds(a.id, b.id))[0];
  const bestCodex = models
    .filter((model) => isOpenAiCodexModelId(model.id))
    .sort((a, b) => compareOpenAiModelIds(a.id, b.id))[0];

  if (bestGeneral && bestCodex) {
    return shouldPreferOpenAiGeneralModel(bestGeneral.id, bestCodex.id) ? bestGeneral : bestCodex;
  }

  if (bestGeneral) return bestGeneral;
  if (bestCodex) return bestCodex;

  return models.slice().sort((a, b) => compareOpenAiModelIds(a.id, b.id))[0] ?? null;
}

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

void test("GPT-5.6 registry exposes the metadata used by the app", () => {
  for (const provider of OPENAI_PROVIDERS) {
    const isCodex = provider === "openai-codex";

    for (const expected of GPT_56_VARIANTS) {
      const model = getModel(provider, expected.id);
      assert.equal(model.id, expected.id);
      assert.equal(model.name, expected.name);
      assert.equal(model.provider, provider);
      assert.equal(model.api, isCodex ? "openai-codex-responses" : "openai-responses");
      assert.equal(
        model.baseUrl,
        isCodex ? "https://chatgpt.com/backend-api" : "https://api.openai.com/v1",
      );
      assert.equal(model.reasoning, true);
      assert.deepEqual(model.input, ["text", "image"]);
      assert.equal(model.contextWindow, 272_000);
      assert.equal(model.maxTokens, 128_000);
      assert.deepEqual(model.cost, isCodex && "codexCost" in expected ? expected.codexCost : expected.cost);
      assert.equal(model.thinkingLevelMap?.minimal, isCodex ? "low" : null);
      assert.equal(model.thinkingLevelMap?.xhigh, "xhigh");
      assert.equal(model.thinkingLevelMap?.max, "max");
      assert.deepEqual(getThinkingLevelsForModel(model), GPT_56_THINKING_LEVELS[provider]);
    }
  }
});

void test("model selector orders all GPT-5.6 tiers as Sol, Terra, Luna", () => {
  const items = GPT_56_IDS
    .slice()
    .reverse()
    .map((id) => ({ provider: "openai-codex", id, model: getModel("openai-codex", id) }));

  const ordered = orderModelsForSelector(items, null);
  assert.deepEqual(ordered.map((item) => item.id), GPT_56_IDS);
});

void test("Claude Fable 5 registry metadata is usable by the add-in", () => {
  const fable = getModel("anthropic", "claude-fable-5");
  assert.equal(fable.provider, "anthropic");
  assert.equal(fable.api, "anthropic-messages");
  assert.ok(fable.reasoning, "expected Fable 5 to support reasoning");
  assert.ok(fable.contextWindow >= 1_000_000, "expected a 1M-token context window");
});

void test("pickDefaultModel prefers the latest Opus for Anthropic-only setups", () => {
  const models = getModels("anthropic");
  const opus = models.filter((m) => m.id.startsWith("claude-opus-"));
  assert.ok(opus.length > 0, "expected at least one Opus model in the registry");

  const selected = pickDefaultModel(["anthropic"]);
  assert.equal(selected.provider, "anthropic");
  assert.equal(selected.id, "claude-opus-5");
});

void test("current OpenAI providers select GPT-5.6 Sol as the default", () => {
  for (const provider of OPENAI_PROVIDERS) {
    const selected = pickDefaultModel([provider]);
    assert.equal(selected.provider, provider);
    assert.equal(selected.id, "gpt-5.6-sol");
  }
});

void test("Bedrock provider uses the browser-safe unsupported-provider stub", async () => {
  installBedrockProviderStub();

  const selected = await completeSimple(
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

void test("pickDefaultModel matches the current OpenAI default-selection contract", () => {
  for (const provider of OPENAI_PROVIDERS) {
    const expected = pickExpectedOpenAiDefault(provider);
    assert.ok(expected, `expected OpenAI default candidate for ${provider}`);

    const selected = pickDefaultModel([provider]);
    assert.equal(selected.provider, provider);
    assert.equal(selected.id, expected.id);
  }
});

void test("pickDefaultModel falls back to the preferred hardcoded OpenAI default", () => {
  const selected = pickDefaultModel([]);
  assert.equal(selected.provider, "openai");
  assert.equal(selected.id, "gpt-5.6-sol");
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
  assert.equal(selected.id, "gpt-5.6-sol");
});

void test("pickDefaultModel prefers GPT-5.6 Sol when OpenAI and Anthropic are both available", () => {
  const selected = pickDefaultModel(["anthropic", "openai"]);
  assert.equal(selected.provider, "openai");
  assert.equal(selected.id, "gpt-5.6-sol");
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
  const runtime: { process?: DynamicValue } = {};
  installProcessEnvShim(runtime);

  assert.ok(runtime.process && typeof runtime.process === "object" && !Array.isArray(runtime.process));

  if (!runtime.process || typeof runtime.process !== "object" || Array.isArray(runtime.process)) {
    assert.fail("expected process shim object");
  }

  assert.equal("env" in runtime.process, true);
  if (!("env" in runtime.process)) {
    assert.fail("expected process.env to exist");
  }

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

void test("vite proxy orders overlapping routes from most specific to least specific", () => {
  const viteConfigPath = path.resolve(process.cwd(), "vite.config.ts");
  const content = readFileSync(viteConfigPath, "utf8");

  const anthropicPlatformIndex = content.indexOf('"/oauth-proxy/anthropic-platform"');
  const anthropicIndex = content.indexOf('"/oauth-proxy/anthropic"');
  const sandboxIndex = content.indexOf('"/api-proxy/google-cloudcode-sandbox"');
  const cloudcodeIndex = content.indexOf('"/api-proxy/google-cloudcode"');
  const googleIndex = content.indexOf('"/api-proxy/google"');

  assert.notEqual(anthropicPlatformIndex, -1, "expected Anthropic platform OAuth route");
  assert.notEqual(anthropicIndex, -1, "expected Anthropic OAuth route");
  assert.notEqual(sandboxIndex, -1, "expected sandbox proxy route");
  assert.notEqual(cloudcodeIndex, -1, "expected cloudcode proxy route");
  assert.notEqual(googleIndex, -1, "expected generic google proxy route");

  assert.ok(
    anthropicPlatformIndex < anthropicIndex,
    "Anthropic platform route must come before generic Anthropic route",
  );
  assert.ok(
    sandboxIndex < cloudcodeIndex,
    "sandbox route must come before cloudcode route",
  );
  assert.ok(
    cloudcodeIndex < googleIndex,
    "cloudcode route must come before generic google route",
  );
});

void test("vite aliases Ajv packages to local stubs for CSP-safe Office builds", () => {
  const viteConfigPath = path.resolve(process.cwd(), "vite.config.ts");
  const content = readFileSync(viteConfigPath, "utf8");

  assert.notEqual(
    content.indexOf("function buildBrowserAliasMap()"),
    -1,
    "expected centralized browser alias helper",
  );
  assert.notEqual(
    content.indexOf('ajv: resolveFromRoot("src/stubs/ajv.ts")'),
    -1,
    "expected Ajv alias to local CSP-safe stub",
  );
  assert.notEqual(
    content.indexOf('"ajv-formats": resolveFromRoot("src/stubs/ajv-formats.ts")'),
    -1,
    "expected ajv-formats alias to local no-op stub",
  );
  assert.notEqual(
    content.indexOf("alias: buildBrowserAliases()"),
    -1,
    "expected resolve.alias to use centralized browser alias helper",
  );
  assert.equal(
    content.indexOf("pi-web-ui"),
    -1,
    "expected no pi-web-ui aliases/stubs to remain in vite.config (UI is first-party)",
  );
});

void test("Ajv stubs keep fallback behavior explicit", () => {
  const ajvStubPath = path.resolve(process.cwd(), "src/stubs/ajv.ts");
  const ajvFormatsStubPath = path.resolve(process.cwd(), "src/stubs/ajv-formats.ts");

  const ajvStubContent = readFileSync(ajvStubPath, "utf8");
  const ajvFormatsStubContent = readFileSync(ajvFormatsStubPath, "utf8");

  assert.notEqual(
    ajvStubContent.indexOf('throw new Error("Ajv disabled: Office Add-in CSP does not allow unsafe-eval")'),
    -1,
    "expected Ajv stub constructor to throw so pi-ai disables schema validation",
  );
  assert.notEqual(
    ajvFormatsStubContent.indexOf("export default function addFormats()"),
    -1,
    "expected ajv-formats stub to expose a no-op default export",
  );
});

void test("vite deduplicates marked so the safety patch covers all instances", () => {
  const viteConfigPath = path.resolve(process.cwd(), "vite.config.ts");
  const content = readFileSync(viteConfigPath, "utf8");

  // The resolve.dedupe config forces all `import ... from "marked"` to
  // resolve to the same module instance, ensuring installMarkedSafetyPatch()
  // intercepts markdown-block's .use() calls (it ships its own marked copy).
  assert.match(
    content,
    /dedupe.*\[.*"marked".*\]/s,
    'expected resolve.dedupe to include "marked" — without it, markdown-block ' +
    "uses a separate marked instance that our KaTeX safety patch never touches",
  );
});
