import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentSkillDefinition } from "../src/skills/types.ts";
import { createSkillReadCache } from "../src/skills/read-cache.ts";
import {
  isSkillsErrorDetails,
  isSkillsInstallDetails,
  isSkillsListDetails,
  isSkillsReadDetails,
  isSkillsUninstallDetails,
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

  assert.ok(isSkillsListDetails(result.details));
  if (!isSkillsListDetails(result.details)) return;

  assert.equal(result.details.count, 1);
  assert.equal(result.details.externalDiscoveryEnabled, false);
  assert.deepEqual(result.details.names, ["web-search"]);
  assert.deepEqual(result.details.entries[0], {
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

  assert.ok(isSkillsReadDetails(first.details));
  assert.ok(isSkillsReadDetails(second.details));
  if (!isSkillsReadDetails(first.details) || !isSkillsReadDetails(second.details)) return;

  assert.equal(first.details.cacheHit, false);
  assert.equal(second.details.cacheHit, true);
  assert.equal(first.details.sourceKind, "bundled");
  assert.equal(second.details.sourceKind, "bundled");
  assert.equal(second.details.location, "skills/web-search/SKILL.md");
  assert.equal(second.details.readCount, 1);
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

  assert.ok(isSkillsReadDetails(refreshed.details));
  if (!isSkillsReadDetails(refreshed.details)) return;

  assert.equal(refreshed.details.cacheHit, false);
  assert.equal(refreshed.details.refreshed, true);
  assert.equal(refreshed.details.readCount, 2);
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

  assert.ok(isSkillsReadDetails(second.details));
  if (!isSkillsReadDetails(second.details)) return;

  assert.equal(second.details.cacheHit, false);
  assert.equal(second.details.readCount, 1);
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
  assert.ok(isSkillsErrorDetails(result.details));
  if (!isSkillsErrorDetails(result.details)) return;

  assert.equal(result.details.externalDiscoveryEnabled, false);
  assert.deepEqual(result.details.availableNames, ["web-search"]);
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

  assert.ok(isSkillsListDetails(result.details));
  if (!isSkillsListDetails(result.details)) return;

  assert.equal(result.details.externalDiscoveryEnabled, true);
  assert.deepEqual(result.details.names, ["custom-skill", "web-search"]);
  assert.equal(result.details.entries.find((entry) => entry.name === "custom-skill")?.sourceKind, "external");
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

  assert.ok(isSkillsReadDetails(result.details));
  if (!isSkillsReadDetails(result.details)) return;

  assert.equal(result.details.sourceKind, "external");
  assert.equal(result.details.location, CUSTOM_EXTERNAL_SKILL.location);
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

  assert.ok(isSkillsListDetails(listResult.details));
  if (!isSkillsListDetails(listResult.details)) return;

  assert.deepEqual(listResult.details.names, ["web-search"]);

  const readResult = await tool.execute("call-read-disabled", {
    action: "read",
    name: "custom-skill",
  });

  assert.ok(isSkillsErrorDetails(readResult.details));
  if (!isSkillsErrorDetails(readResult.details)) return;

  assert.match(readResult.details.message, /Skill not found: `custom-skill`/);
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

  assert.ok(isSkillsReadDetails(firstRead.details));
  if (!isSkillsReadDetails(firstRead.details)) return;

  assert.equal(firstRead.details.cacheHit, false);

  disabled = true;

  const secondRead = await tool.execute("call-read-disabled-after-cache", {
    action: "read",
    name: "web-search",
  });

  assert.ok(isSkillsErrorDetails(secondRead.details));
  if (!isSkillsErrorDetails(secondRead.details)) return;

  assert.match(secondRead.details.message, /Skill not found: `web-search`/);
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

  assert.ok(isSkillsInstallDetails(result.details));
  if (!isSkillsInstallDetails(result.details)) return;

  assert.equal(result.details.skillName, "custom-skill");
  assert.equal(result.details.location, "skills/external/custom-skill/SKILL.md");
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

  assert.ok(isSkillsErrorDetails(result.details));
  if (!isSkillsErrorDetails(result.details)) return;

  assert.equal(result.details.action, "install");
  assert.match(result.details.message, /markdown is required/i);
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

  assert.ok(isSkillsUninstallDetails(removed.details));
  if (!isSkillsUninstallDetails(removed.details)) return;

  assert.equal(removed.details.skillName, "custom-skill");
  assert.equal(removed.details.removed, true);

  const missing = await tool.execute("call-uninstall-no", {
    action: "uninstall",
    name: "missing-skill",
  });

  assert.ok(isSkillsUninstallDetails(missing.details));
  if (!isSkillsUninstallDetails(missing.details)) return;

  assert.equal(missing.details.removed, false);
  assert.equal(changedCount, 1);
});
