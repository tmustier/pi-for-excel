import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentSkillDefinition } from "../src/skills/types.ts";
import { createSkillReadCache } from "../src/skills/read-cache.ts";
import {
  decodeToolDetailsOfKind,
} from "../src/tools/tool-details.ts";
import { createSkillsTool } from "../src/tools/skills.ts";

const WEB_SEARCH_SKILL: AgentSkillDefinition = {
  name: "web-search",
  description: "Search the web for fresh facts.",
  compatibility: "Requires web_search integration.",
  location: "skills/web-search/SKILL.md",
  sourceKind: "bundled",
  markdown: "# Web Search\n\nUse web search when workbook context is insufficient.",
  body: "# Web Search\n\nUse web search when workbook context is insufficient.",
};

const CUSTOM_EXTERNAL_SKILL: AgentSkillDefinition = {
  name: "custom-skill",
  description: "Custom external skill.",
  compatibility: "External discovery test",
  location: "/Users/test/.pi/skills/custom-skill/SKILL.md",
  sourceKind: "external",
  markdown: "# Custom Skill\n\nExternal skill body.",
  body: "# Custom Skill\n\nExternal skill body.",
};

void test("skills list renders provenance and structured list details", async () => {
  const tool = createSkillsTool({
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => false,
    loadExternalSkills: () => Promise.resolve([]),
  });

  const result = await tool.execute("call-list", { action: "list" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /Available Agent Skills \(1\)/);
  assert.match(text, /source: bundled/i);

  const resultDetails = decodeToolDetailsOfKind(result.details, "skills_list");
  assert.ok(resultDetails);

  assert.equal(resultDetails.count, 1);
  assert.equal(resultDetails.externalDiscoveryEnabled, false);
  assert.deepEqual(resultDetails.names, ["web-search"]);
  assert.deepEqual(resultDetails.entries[0], {
    name: "web-search",
    sourceKind: "bundled",
    location: "skills/web-search/SKILL.md",
  });
});

void test("skills read uses session cache and reports cacheHit details", async () => {
  const cache = createSkillReadCache();

  const tool = createSkillsTool({
    getSessionId: () => "session-1",
    readCache: cache,
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => false,
    loadExternalSkills: () => Promise.resolve([]),
  });

  const first = await tool.execute("call-read-1", { action: "read", name: "web-search" });
  const second = await tool.execute("call-read-2", { action: "read", name: "web-search" });

  const firstDetails = decodeToolDetailsOfKind(first.details, "skills_read");
  const secondDetails = decodeToolDetailsOfKind(second.details, "skills_read");
  assert.ok(firstDetails);
  assert.ok(secondDetails);

  assert.equal(firstDetails.cacheHit, false);
  assert.equal(secondDetails.cacheHit, true);
  assert.equal(firstDetails.sourceKind, "bundled");
  assert.equal(secondDetails.sourceKind, "bundled");
  assert.equal(secondDetails.location, "skills/web-search/SKILL.md");
  assert.equal(secondDetails.readCount, 1);
});

void test("skills read with refresh=true bypasses cache and reports refreshed details", async () => {
  const cache = createSkillReadCache();

  const tool = createSkillsTool({
    getSessionId: () => "session-refresh",
    readCache: cache,
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => false,
    loadExternalSkills: () => Promise.resolve([]),
  });

  await tool.execute("call-read-1", { action: "read", name: "web-search" });
  const refreshed = await tool.execute("call-read-2", {
    action: "read",
    name: "web-search",
    refresh: true,
  });

  const refreshedDetails = decodeToolDetailsOfKind(refreshed.details, "skills_read");
  assert.ok(refreshedDetails);

  assert.equal(refreshedDetails.cacheHit, false);
  assert.equal(refreshedDetails.refreshed, true);
  assert.equal(refreshedDetails.readCount, 2);
});

void test("skills read cache is session-scoped", async () => {
  let currentSession = "session-a";

  const tool = createSkillsTool({
    getSessionId: () => currentSession,
    readCache: createSkillReadCache(),
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => false,
    loadExternalSkills: () => Promise.resolve([]),
  });

  await tool.execute("call-read-a", { action: "read", name: "web-search" });
  currentSession = "session-b";
  const second = await tool.execute("call-read-b", { action: "read", name: "web-search" });

  const secondDetails = decodeToolDetailsOfKind(second.details, "skills_read");
  assert.ok(secondDetails);

  assert.equal(secondDetails.cacheHit, false);
  assert.equal(secondDetails.readCount, 1);
});

void test("skills read without name returns structured error details", async () => {
  const tool = createSkillsTool({
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => false,
    loadExternalSkills: () => Promise.resolve([]),
  });

  const result = await tool.execute("call-err", { action: "read" });
  const text = result.content[0]?.type === "text" ? result.content[0].text : "";

  assert.match(text, /name is required/i);
  const resultDetails = decodeToolDetailsOfKind(result.details, "skills_error");
  assert.ok(resultDetails);

  assert.equal(resultDetails.externalDiscoveryEnabled, false);
  assert.deepEqual(resultDetails.availableNames, ["web-search"]);
});

void test("skills list includes external entries when discovery is enabled", async () => {
  const tool = createSkillsTool({
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => true,
    loadExternalSkills: () => Promise.resolve([CUSTOM_EXTERNAL_SKILL]),
  });

  const result = await tool.execute("call-list-ext", { action: "list" });

  const resultDetails = decodeToolDetailsOfKind(result.details, "skills_list");
  assert.ok(resultDetails);

  assert.equal(resultDetails.externalDiscoveryEnabled, true);
  assert.deepEqual(resultDetails.names, ["custom-skill", "web-search"]);
  assert.equal(resultDetails.entries.find((entry) => entry.name === "custom-skill")?.sourceKind, "external");
});

void test("skills read resolves external skill when discovery is enabled", async () => {
  const tool = createSkillsTool({
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => true,
    loadExternalSkills: () => Promise.resolve([CUSTOM_EXTERNAL_SKILL]),
  });

  const result = await tool.execute("call-read-ext", { action: "read", name: "custom-skill" });

  const text = result.content[0]?.type === "text" ? result.content[0].text : "";
  assert.match(text, /Custom Skill/);

  const resultDetails = decodeToolDetailsOfKind(result.details, "skills_read");
  assert.ok(resultDetails);

  assert.equal(resultDetails.sourceKind, "external");
  assert.equal(resultDetails.location, CUSTOM_EXTERNAL_SKILL.location);
});

void test("skills tool exposes no skills when activation state is unreadable", async () => {
  const previousWindow = Reflect.get(globalThis, "window");
  Reflect.set(globalThis, "window", {});

  try {
    const tool = createSkillsTool({
      catalog: {
        list: () => [WEB_SEARCH_SKILL],
      },
      isExternalDiscoveryEnabled: () => false,
      loadExternalSkills: () => Promise.resolve([]),
    });

    await assert.rejects(
      tool.execute("call-list-unreadable-activation", { action: "list" }),
      /AppStorage not initialized/u,
    );
  } finally {
    if (previousWindow === undefined) {
      Reflect.deleteProperty(globalThis, "window");
    } else {
      Reflect.set(globalThis, "window", previousWindow);
    }
  }
});

void test("skills list/read exclude disabled skills", async () => {
  const tool = createSkillsTool({
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => true,
    loadExternalSkills: () => Promise.resolve([CUSTOM_EXTERNAL_SKILL]),
    loadDisabledSkillNames: () => Promise.resolve(new Set(["custom-skill"])),
  });

  const listResult = await tool.execute("call-list-disabled", { action: "list" });

  const listResultDetails = decodeToolDetailsOfKind(listResult.details, "skills_list");
  assert.ok(listResultDetails);

  assert.deepEqual(listResultDetails.names, ["web-search"]);

  const readResult = await tool.execute("call-read-disabled", {
    action: "read",
    name: "custom-skill",
  });

  const readResultDetails = decodeToolDetailsOfKind(readResult.details, "skills_error");
  assert.ok(readResultDetails);

  assert.match(readResultDetails.message, /Skill not found: `custom-skill`/);
});

void test("skills read ignores stale cache entries when skill becomes disabled", async () => {
  const cache = createSkillReadCache();
  let disabled = false;

  const tool = createSkillsTool({
    getSessionId: () => "session-disable",
    readCache: cache,
    catalog: {
      list: () => [WEB_SEARCH_SKILL],
    },
    isExternalDiscoveryEnabled: () => false,
    loadExternalSkills: () => Promise.resolve([]),
    loadDisabledSkillNames: () => Promise.resolve(disabled ? new Set(["web-search"]) : new Set()),
  });

  const firstRead = await tool.execute("call-read-enabled", {
    action: "read",
    name: "web-search",
  });

  const firstReadDetails = decodeToolDetailsOfKind(firstRead.details, "skills_read");
  assert.ok(firstReadDetails);

  assert.equal(firstReadDetails.cacheHit, false);

  disabled = true;

  const secondRead = await tool.execute("call-read-disabled-after-cache", {
    action: "read",
    name: "web-search",
  });

  const secondReadDetails = decodeToolDetailsOfKind(secondRead.details, "skills_error");
  assert.ok(secondReadDetails);

  assert.match(secondReadDetails.message, /Skill not found: `web-search`/);
});

void test("skills install writes external skill and emits structured install details", async () => {
  let installedName = "";
  let installedMarkdown = "";
  let changedReason = "";

  const tool = createSkillsTool({
    catalog: { list: () => [WEB_SEARCH_SKILL] },
    isExternalDiscoveryEnabled: () => true,
    loadExternalSkills: () => Promise.resolve([]),
    installExternalSkill: (args) => {
      installedName = args.name;
      installedMarkdown = args.markdown;
      return Promise.resolve({
        name: args.name,
        location: `skills/external/${args.name}/SKILL.md`,
      });
    },
    dispatchSkillsChanged: (reason) => {
      changedReason = reason;
    },
  });

  const markdown = "---\nname: custom-skill\ndescription: Custom\n---\n\nBody";
  const result = await tool.execute("call-install", {
    action: "install",
    name: "custom-skill",
    markdown,
  });

  assert.equal(installedName, "custom-skill");
  assert.equal(installedMarkdown, markdown);
  assert.equal(changedReason, "catalog");

  const resultDetails = decodeToolDetailsOfKind(result.details, "skills_install");
  assert.ok(resultDetails);

  assert.equal(resultDetails.skillName, "custom-skill");
  assert.equal(resultDetails.location, "skills/external/custom-skill/SKILL.md");
});

void test("skills install requires markdown", async () => {
  const tool = createSkillsTool({
    catalog: { list: () => [WEB_SEARCH_SKILL] },
    isExternalDiscoveryEnabled: () => false,
    loadExternalSkills: () => Promise.resolve([]),
  });

  const result = await tool.execute("call-install-missing", {
    action: "install",
    name: "custom-skill",
  });

  const resultDetails = decodeToolDetailsOfKind(result.details, "skills_error");
  assert.ok(resultDetails);

  assert.equal(resultDetails.action, "install");
  assert.match(resultDetails.message, /markdown is required/i);
});

void test("skills uninstall reports removed state and emits refresh only when removed", async () => {
  let changedCount = 0;

  const tool = createSkillsTool({
    catalog: { list: () => [WEB_SEARCH_SKILL] },
    isExternalDiscoveryEnabled: () => true,
    loadExternalSkills: () => Promise.resolve([]),
    uninstallExternalSkill: (args) => Promise.resolve(args.name === "custom-skill"),
    dispatchSkillsChanged: () => {
      changedCount += 1;
    },
  });

  const removed = await tool.execute("call-uninstall-yes", {
    action: "uninstall",
    name: "custom-skill",
  });

  const removedDetails = decodeToolDetailsOfKind(removed.details, "skills_uninstall");
  assert.ok(removedDetails);

  assert.equal(removedDetails.skillName, "custom-skill");
  assert.equal(removedDetails.removed, true);

  const missing = await tool.execute("call-uninstall-no", {
    action: "uninstall",
    name: "missing-skill",
  });

  const missingDetails = decodeToolDetailsOfKind(missing.details, "skills_uninstall");
  assert.ok(missingDetails);

  assert.equal(missingDetails.removed, false);
  assert.equal(changedCount, 1);
});
