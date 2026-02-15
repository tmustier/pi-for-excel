import assert from "node:assert/strict";
import { test } from "node:test";

import {
  compareModels,
  modelRecencyScore,
  parseMajorMinor,
  providerPriority,
} from "../src/models/model-ordering.ts";
import { BROWSER_OAUTH_PROVIDERS, mapToApiProvider } from "../src/auth/provider-map.ts";

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

void test("browser oauth providers include openai-codex", () => {
  assert.equal(BROWSER_OAUTH_PROVIDERS.includes("openai-codex"), true);
});
