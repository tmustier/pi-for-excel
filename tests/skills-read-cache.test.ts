import assert from "node:assert/strict";
import { test } from "node:test";

import { createSkillReadCache } from "../src/skills/read-cache.ts";

void test("skill read cache stores and retrieves per session", () => {
  const cache = createSkillReadCache();
  const entry = cache.set("session-1", {
    skillName: "web-search",
    sourceKind: "bundled",
    location: "skills/web-search/SKILL.md",
    markdown: "# Web Search",
  });

  assert.equal(entry.readCount, 1);

  const cached = cache.get("session-1", "web-search");
  assert.ok(cached);
  assert.equal(cached.skillName, "web-search");
  assert.equal(cached.sourceKind, "bundled");
  assert.equal(cached.location, "skills/web-search/SKILL.md");
  assert.equal(cached.markdown, "# Web Search");
  assert.equal(cached.readCount, 1);
});

void test("skill read cache updates readCount and supports session invalidation", () => {
  const cache = createSkillReadCache();
  cache.set("session-1", {
    skillName: "web-search",
    sourceKind: "bundled",
    location: "skills/web-search/SKILL.md",
    markdown: "# Web Search",
  });
  const updated = cache.set("session-1", {
    skillName: "web-search",
    sourceKind: "bundled",
    location: "skills/web-search/SKILL.md",
    markdown: "# Web Search v2",
  });

  assert.equal(updated.readCount, 2);
  assert.equal(cache.get("session-1", "web-search")?.markdown, "# Web Search v2");

  cache.clearSession("session-1");
  assert.equal(cache.get("session-1", "web-search"), null);
});
