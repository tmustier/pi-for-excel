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
  modelRecencyScore,
  openAiFamilyPriority,
  openAiVariantPriority,
  parseMajorMinor,
  providerPriority,
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
      input: 2.5,
      output: 15,
      cacheRead: 0.25,
      cacheWrite: 3.125,
      tiers: [{ inputTokensAbove: 272_000, input: 5, output: 22.5, cacheRead: 0.5, cacheWrite: 6.25 }],
    },
  },
  {
    id: "gpt-5.6-luna",
    name: "GPT-5.6 Luna",
    cost: {
      input: 1,
      output: 6,
      cacheRead: 0.1,
      cacheWrite: 1.25,
      tiers: [{ inputTokensAbove: 272_000, input: 2, output: 9, cacheRead: 0.2, cacheWrite: 2.5 }],
    },
  },
];

const GPT_56_IDS = GPT_56_VARIANTS.map((variant) => variant.id);
const GPT_56_THINKING_LEVELS = ["off", "minimal", "low", "medium", "high", "xhigh", "max"];

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

void test("parseMajorMinor packs Claude-style -major-minor as major*10+minor", () => {
  assert.equal(parseMajorMinor("claude-opus-4-6"), 46);
  assert.equal(parseMajorMinor("claude-opus-4-7"), 47);
});

void test("parseMajorMinor scores Claude Fable ids as a new major version", () => {
  assert.equal(parseMajorMinor("claude-fable-5"), 50);
  assert.equal(parseMajorMinor("anthropic/claude-fable-5"), 50);
  assert.equal(parseMajorMinor("us.anthropic.claude-fable-5"), 50);
});

void test("parseMajorMinor handles namespaced and dotted Claude registry ids", () => {
  assert.equal(parseMajorMinor("claude-opus-4.7"), 47);
  assert.equal(parseMajorMinor("anthropic.claude-opus-4-1-20250805-v1:0"), 41);
  assert.equal(parseMajorMinor("eu.anthropic.claude-sonnet-4-6"), 46);
  assert.equal(parseMajorMinor("anthropic/claude-opus-4.7"), 47);
});

void test("parseMajorMinor does not treat YYYYMMDD as a minor version", () => {
  // This used to incorrectly parse as 4.20250514.
  assert.equal(parseMajorMinor("claude-opus-4-20250514"), 40);
});

void test("parseMajorMinor handles dot-style versions", () => {
  assert.equal(parseMajorMinor("gpt-5.3-codex"), 53);
  assert.equal(parseMajorMinor("gpt-5.5"), 55);
  assert.equal(parseMajorMinor("gpt-5.6-sol"), 56);
  assert.equal(parseMajorMinor("gpt-5.6-terra"), 56);
  assert.equal(parseMajorMinor("gpt-5.6-luna"), 56);
  assert.equal(parseMajorMinor("gemini-2.5-pro"), 25);
  assert.equal(parseMajorMinor("google/gemini-3.1-pro-preview"), 31);
});

void test("parseMajorMinor ignores OpenAI dated suffixes after the family name", () => {
  assert.equal(parseMajorMinor("gpt-4o-2024-11-20"), 40);
});

void test("parseMajorMinor ignores Gemini preview date suffixes", () => {
  assert.equal(parseMajorMinor("gemini-2.5-pro-preview-06-05"), 25);
});

void test("parseMajorMinor supports 2-digit minors (e.g. 5.12)", () => {
  assert.equal(parseMajorMinor("gpt-5.12"), 512);
});

void test("parseMajorMinor falls back for non-Claude/GPT/Gemini registry families", () => {
  assert.equal(parseMajorMinor("gemma-4-31b-it"), 40);
  assert.equal(parseMajorMinor("gemma-3-27b-it"), 30);
  assert.equal(parseMajorMinor("zai.glm-5"), 50);
  assert.equal(parseMajorMinor("zai.glm-4.7"), 47);
  assert.equal(parseMajorMinor("deepseek.v3.2"), 32);
  assert.equal(parseMajorMinor("MiniMax-M2.7"), 27);
  assert.equal(parseMajorMinor("moonshotai/Kimi-K2.6"), 26);
  assert.equal(parseMajorMinor("Qwen/Qwen3.5-coder"), 35);
  assert.equal(parseMajorMinor("amazon.nova-2-lite-v1:0"), 20);
  assert.equal(parseMajorMinor("unknown-model-2024-11-20"), 0);
});

void test("OpenAI model helpers handle GPT-5.6 tiers and Codex variants", () => {
  assert.equal(isOpenAiGeneralGptModelId("gpt-5"), true);
  assert.equal(isOpenAiGeneralGptModelId("gpt-5-pro"), true);
  assert.equal(isOpenAiGeneralGptModelId("gpt-5.6-sol"), true);
  assert.equal(isOpenAiGeneralGptModelId("gpt-5.6-terra"), true);
  assert.equal(isOpenAiGeneralGptModelId("gpt-5.6-luna"), true);
  assert.equal(isOpenAiGeneralGptModelId("gpt-5-codex"), false);
  assert.equal(isOpenAiCodexModelId("gpt-5-codex"), true);
  assert.equal(isOpenAiCodexModelId("gpt-5.1-codex-max"), true);
});

void test("openAiFamilyPriority prefers base GPT-5 over Codex variants", () => {
  assert.ok(openAiFamilyPriority("gpt-5") < openAiFamilyPriority("gpt-5-pro"));
  assert.ok(openAiFamilyPriority("gpt-5-pro") < openAiFamilyPriority("gpt-5-codex"));
});

void test("compareOpenAiModelIds prefers newer versions before family tie-breaks", () => {
  const ids = ["gpt-4o-2024-11-20", "gpt-5.5", "gpt-5.4-pro", "gpt-5.3-codex"];
  ids.sort(compareOpenAiModelIds);
  assert.deepEqual(ids, ["gpt-5.5", "gpt-5.4-pro", "gpt-5.3-codex", "gpt-4o-2024-11-20"]);
});

void test("compareOpenAiModelIds prefers plain GPT-5 over other same-version variants", () => {
  const ids = ["gpt-5-pro", "gpt-5-codex", "gpt-5"];
  ids.sort(compareOpenAiModelIds);
  assert.deepEqual(ids, ["gpt-5", "gpt-5-pro", "gpt-5-codex"]);
});

void test("compareOpenAiModelIds orders GPT-5.6 Sol, Terra, then Luna", () => {
  const ids = ["gpt-5.6-luna", "gpt-5.6-sol", "gpt-5.6-terra"];
  ids.sort(compareOpenAiModelIds);
  assert.deepEqual(ids, GPT_56_IDS);
  assert.ok(openAiVariantPriority("gpt-5.6-sol") < openAiVariantPriority("gpt-5.6-terra"));
  assert.ok(openAiVariantPriority("gpt-5.6-terra") < openAiVariantPriority("gpt-5.6-luna"));
});

void test("shouldPreferOpenAiGeneralModel only prefers GPT when it is as new or newer", () => {
  assert.equal(shouldPreferOpenAiGeneralModel("gpt-5", "gpt-5-codex"), true);
  assert.equal(shouldPreferOpenAiGeneralModel("gpt-5.5", "gpt-5.3-codex"), true);
  assert.equal(shouldPreferOpenAiGeneralModel("gpt-5-pro", "gpt-5.1-codex-max"), false);
});

void test("current Pi registry exposes only the three canonical GPT-5.6 tier IDs", () => {
  for (const provider of OPENAI_PROVIDERS) {
    const ids = getModels(provider)
      .filter((model) => model.id.startsWith("gpt-5.6"))
      .map((model) => model.id)
      .sort();

    assert.deepEqual(ids, GPT_56_IDS.slice().sort(), `unexpected GPT-5.6 IDs for ${provider}`);
    assert.equal(modelsRuntime.getModel(provider, "gpt-5.6"), undefined, `bare alias must not exist for ${provider}`);
  }

  assert.equal(getModel("anthropic", "claude-opus-4-7").id, "claude-opus-4-7");
  assert.equal(getModel("anthropic", "claude-fable-5").id, "claude-fable-5");
  assert.equal(getModel("google", "gemini-3.1-pro-preview").id, "gemini-3.1-pro-preview");
});

void test("GPT-5.6 registry metadata exactly matches native Pi 0.80.8", () => {
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
      assert.equal(model.contextWindow, isCodex ? 372_000 : 272_000);
      assert.equal(model.maxTokens, 128_000);
      assert.deepEqual(model.cost, expected.cost);
      assert.deepEqual(
        model.thinkingLevelMap,
        isCodex
          ? { xhigh: "xhigh", max: "max", minimal: "low" }
          : { off: "none", xhigh: "xhigh", max: "max" },
      );
      assert.deepEqual(getThinkingLevelsForModel(model), GPT_56_THINKING_LEVELS);
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
  assert.equal(selected.id, "claude-opus-4-8");
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

void test("modelRecencyScore prefers higher version, then later date suffix", () => {
  assert.ok(
    modelRecencyScore("claude-opus-4-20250201") > modelRecencyScore("claude-opus-4-20250101"),
    "expected 20250201 > 20250101 for same major",
  );

  assert.ok(
    modelRecencyScore("gpt-4o-2024-11-20") > modelRecencyScore("gpt-4o-2024-05-13"),
    "expected 2024-11-20 > 2024-05-13 for dated GPT snapshots",
  );

  assert.ok(
    modelRecencyScore("anthropic.claude-opus-4-1-20250805-v1:0") >
      modelRecencyScore("anthropic.claude-opus-4-1-20250101-v1:0"),
    "expected embedded Bedrock compact dates to affect recency",
  );

  assert.ok(
    modelRecencyScore("gemini-2.5-flash-preview-05-20") > modelRecencyScore("gemini-2.5-flash-preview-04-17"),
    "expected 05-20 > 04-17 for dated Gemini previews",
  );

  // Version beats date.
  assert.ok(
    modelRecencyScore("claude-opus-4-6") > modelRecencyScore("claude-opus-4-20250201"),
    "expected 4-6 to outrank 4-YYYYMMDD",
  );
});

void test("compareModels sorts generic versioned registry ids by parsed recency", () => {
  const gemmaModels = [
    { provider: "opencode", id: "gemma-3-27b-it" },
    { provider: "opencode", id: "gemma-4-31b-it" },
  ];
  gemmaModels.sort(compareModels);
  assert.deepEqual(gemmaModels.map((m) => m.id), ["gemma-4-31b-it", "gemma-3-27b-it"]);

  const zaiModels = [
    { provider: "openrouter", id: "zai.glm-4.7" },
    { provider: "openrouter", id: "zai.glm-5" },
  ];
  zaiModels.sort(compareModels);
  assert.deepEqual(zaiModels.map((m) => m.id), ["zai.glm-5", "zai.glm-4.7"]);

  const letterPrefixedModels = [
    { provider: "openrouter", id: "MiniMax-M2.7" },
    { provider: "openrouter", id: "MiniMax-M2.6" },
  ];
  letterPrefixedModels.sort(compareModels);
  assert.deepEqual(letterPrefixedModels.map((m) => m.id), ["MiniMax-M2.7", "MiniMax-M2.6"]);
});

void test("compareModels sorts real namespaced registry ids by parsed recency", () => {
  const copilotModels = [
    { provider: "github-copilot", id: "claude-opus-4.5" },
    { provider: "github-copilot", id: "claude-opus-4.7" },
    { provider: "github-copilot", id: "claude-opus-4.6" },
  ];
  copilotModels.sort(compareModels);
  assert.deepEqual(
    copilotModels.map((m) => m.id),
    ["claude-opus-4.7", "claude-opus-4.6", "claude-opus-4.5"],
  );

  const bedrockModels = [
    { provider: "amazon-bedrock", id: "google.gemma-3-27b-it" },
    { provider: "amazon-bedrock", id: "anthropic.claude-opus-4-20250514-v1:0" },
    { provider: "amazon-bedrock", id: "anthropic.claude-opus-4-1-20250805-v1:0" },
    { provider: "amazon-bedrock", id: "anthropic.claude-opus-4-7" },
  ];
  bedrockModels.sort(compareModels);
  assert.deepEqual(
    bedrockModels.map((m) => m.id),
    [
      "anthropic.claude-opus-4-7",
      "anthropic.claude-opus-4-1-20250805-v1:0",
      "anthropic.claude-opus-4-20250514-v1:0",
      "google.gemma-3-27b-it",
    ],
  );
});

void test("compareModels sorts non-OpenAI models by provider, family, then recency", () => {
  const models = [
    { provider: "openai", id: "gpt-5.5" },
    { provider: "anthropic", id: "claude-opus-4-7" },
    { provider: "anthropic", id: "claude-sonnet-4-6" },
    { provider: "anthropic", id: "claude-opus-4-6" },
    { provider: "google", id: "gemini-2.5-pro" },
  ];

  models.sort(compareModels);

  // Provider priority: anthropic first, then openai, then google.
  assert.equal(models[0].provider, "anthropic");

  const last = models.at(-1);
  assert.ok(last);
  assert.equal(last.provider, "google");

  // Within anthropic: opus family first; within opus: 4-7 before 4-6.
  const anthropic = models.filter((m) => m.provider === "anthropic");
  assert.deepEqual(
    anthropic.map((m) => m.id),
    ["claude-opus-4-7", "claude-opus-4-6", "claude-sonnet-4-6"],
  );

  // Sanity: providerPriority is stable
  assert.ok(providerPriority("anthropic") < providerPriority("openai"));
});

void test("compareModels puts the Fable family before Opus/Sonnet/Haiku for Anthropic", () => {
  const models = [
    { provider: "anthropic", id: "claude-haiku-4-5" },
    { provider: "anthropic", id: "claude-opus-4-8" },
    { provider: "anthropic", id: "claude-fable-5" },
    { provider: "anthropic", id: "claude-sonnet-4-6" },
  ];

  models.sort(compareModels);

  assert.deepEqual(
    models.map((m) => m.id),
    ["claude-fable-5", "claude-opus-4-8", "claude-sonnet-4-6", "claude-haiku-4-5"],
  );
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
