import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import { test } from "node:test";

import {
  compareModels,
  modelRecencyScore,
  parseMajorMinor,
  providerPriority,
} from "../src/models/model-ordering.ts";
import { BROWSER_OAUTH_PROVIDERS, mapToApiProvider } from "../src/auth/provider-map.ts";
import { rewriteDevProxyUrl } from "../src/auth/dev-rewrites.ts";
import { installProcessEnvShim } from "../src/compat/process-env-shim.ts";

void test("parseMajorMinor packs Claude-style -major-minor as major*10+minor", () => {
  assert.equal(parseMajorMinor("claude-opus-4-5"), 45);
  assert.equal(parseMajorMinor("claude-opus-4-6"), 46);
});

void test("parseMajorMinor does not treat YYYYMMDD as a minor version", () => {
  // This used to incorrectly parse as 4.20250514.
  assert.equal(parseMajorMinor("claude-opus-4-20250514"), 40);
});

void test("parseMajorMinor handles dot-style versions", () => {
  assert.equal(parseMajorMinor("gpt-5.3-codex"), 53);
  assert.equal(parseMajorMinor("gemini-2.5-pro"), 25);
});

void test("parseMajorMinor supports 2-digit minors (e.g. 5.12)", () => {
  assert.equal(parseMajorMinor("gpt-5.12"), 512);
});

void test("modelRecencyScore prefers higher version, then later date suffix", () => {
  assert.ok(
    modelRecencyScore("claude-opus-4-20250201") > modelRecencyScore("claude-opus-4-20250101"),
    "expected 20250201 > 20250101 for same major",
  );

  // Version beats date.
  assert.ok(
    modelRecencyScore("claude-opus-4-6") > modelRecencyScore("claude-opus-4-20250201"),
    "expected 4-6 to outrank 4-YYYYMMDD",
  );
});

void test("compareModels sorts by provider, family, then recency", () => {
  const models = [
    { provider: "openai", id: "gpt-5.3" },
    { provider: "anthropic", id: "claude-opus-4-6" },
    { provider: "anthropic", id: "claude-sonnet-4-6" },
    { provider: "anthropic", id: "claude-opus-4-5" },
    { provider: "google", id: "gemini-2.5-pro" },
  ];

  models.sort(compareModels);

  // Provider priority: anthropic first, then openai, then google.
  assert.equal(models[0].provider, "anthropic");

  const last = models.at(-1);
  assert.ok(last);
  assert.equal(last.provider, "google");

  // Within anthropic: opus family first; within opus: 4-6 before 4-5.
  const anthropic = models.filter((m) => m.provider === "anthropic");
  assert.deepEqual(
    anthropic.map((m) => m.id),
    ["claude-opus-4-6", "claude-opus-4-5", "claude-sonnet-4-6"],
  );

  // Sanity: providerPriority is stable
  assert.ok(providerPriority("anthropic") < providerPriority("openai"));
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

void test("vite proxy orders Google routes from most specific to least specific", () => {
  const viteConfigPath = path.resolve(process.cwd(), "vite.config.ts");
  const content = readFileSync(viteConfigPath, "utf8");

  const sandboxIndex = content.indexOf('"/api-proxy/google-cloudcode-sandbox"');
  const cloudcodeIndex = content.indexOf('"/api-proxy/google-cloudcode"');
  const googleIndex = content.indexOf('"/api-proxy/google"');

  assert.notEqual(sandboxIndex, -1, "expected sandbox proxy route");
  assert.notEqual(cloudcodeIndex, -1, "expected cloudcode proxy route");
  assert.notEqual(googleIndex, -1, "expected generic google proxy route");

  assert.ok(
    sandboxIndex < cloudcodeIndex,
    "sandbox route must come before cloudcode route",
  );
  assert.ok(
    cloudcodeIndex < googleIndex,
    "cloudcode route must come before generic google route",
  );
});
