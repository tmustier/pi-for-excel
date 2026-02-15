import assert from "node:assert/strict";
import test from "node:test";

import { BROWSER_OAUTH_PROVIDERS, mapToApiProvider } from "../src/auth/provider-map.js";

void test("mapToApiProvider keeps openai-codex separate from openai", () => {
  assert.equal(mapToApiProvider("openai-codex"), "openai-codex");
  assert.equal(mapToApiProvider("openai"), "openai");
});

void test("browser OAuth providers include openai-codex", () => {
  assert.equal(BROWSER_OAUTH_PROVIDERS.includes("openai-codex"), true);
});
