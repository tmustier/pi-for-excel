import assert from "node:assert/strict";
import { test } from "node:test";

import { buildSystemPrompt } from "../src/prompt/system-prompt.ts";
import { resolveConventions } from "../src/conventions/store.ts";
import type { LocalServiceEntry } from "../src/tools/bridge-health.ts";

void test("system prompt includes default placeholders when instructions are absent", () => {
  const prompt = buildSystemPrompt();

  assert.match(prompt, /## Rules/);
  assert.match(prompt, /### All my files/);
  assert.match(prompt, /### This file/);
  assert.match(prompt, /\(No rules set\.\)/);
});

void test("system prompt defaults to YOLO execution mode guidance", () => {
  const prompt = buildSystemPrompt();

  assert.match(prompt, /## Execution mode/);
  assert.match(prompt, /Current mode:\s*\*\*Auto\*\*/);
  assert.match(prompt, /low-friction execution/i);
});

void test("system prompt renders Confirm execution mode guidance", () => {
  const prompt = buildSystemPrompt({ executionMode: "safe" });

  assert.match(prompt, /Current mode:\s*\*\*Confirm\*\*/);
  assert.match(prompt, /explicit user confirmation before mutating workbook tools/i);
  assert.match(prompt, /destructive structure operations as high-risk/i);
});

void test("system prompt embeds provided user and workbook instructions", () => {
  const prompt = buildSystemPrompt({
    userInstructions: "Always use EUR",
    workbookInstructions: "Summary sheet is read-only",
  });

  assert.match(prompt, /Always use EUR/);
  assert.match(prompt, /Summary sheet is read-only/);
  assert.match(prompt, /\*\*instructions\*\* tool/);
});

void test("system prompt omits convention overrides when all defaults", () => {
  const conventions = resolveConventions({});
  const prompt = buildSystemPrompt({ conventions });

  assert.ok(!prompt.includes("Active convention overrides"));
});

void test("system prompt includes convention overrides when customized", () => {
  const conventions = resolveConventions({
    visualDefaults: {
      fontName: "Calibri",
    },
    colorConventions: {
      hardcodedValueColor: "#FF0000",
    },
    customPresets: {
      bps: {
        format: '#,##0 "bps"',
        description: "Basis points",
      },
    },
  });
  const prompt = buildSystemPrompt({ conventions });

  assert.match(prompt, /Custom format presets/);
  assert.match(prompt, /`bps` — Basis points/);
  assert.match(prompt, /Active convention overrides/);
  assert.match(prompt, /Default font: Calibri/);
  assert.match(prompt, /Hardcoded value font color: #FF0000/);
});

void test("system prompt lists the conventions tool", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*conventions\*\*/);
});

void test("system prompt includes workbook history recovery tool", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*workbook_history\*\*/);
  assert.match(prompt, /automatic backups/i);
  assert.match(prompt, /write_cells/);
  assert.match(prompt, /fill_formula/);
  assert.match(prompt, /python_transform_range/);
  assert.match(prompt, /format_cells/);
  assert.match(prompt, /conditional_format/);
  assert.match(prompt, /comments/);
  assert.match(prompt, /modify_structure/);
});

void test("system prompt documents Python tools and Pyodide default (bridge awareness moved to Local Services)", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /### Python/);
  assert.match(prompt, /\*\*python_run\*\*/);
  assert.match(prompt, /\*\*python_transform_range\*\*/);
  assert.match(prompt, /Pyodide/);
  assert.match(prompt, /no setup required/i);
  // Bridge-awareness content moved to dynamic ## Local Services section
  assert.ok(!prompt.includes("native Python bridge"), "TOOLS section should not contain bridge-awareness content");
});

void test("system prompt documents trace_dependencies precedents/dependents modes", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*trace_dependencies\*\*/);
  assert.match(prompt, /mode:\s*`precedents`/i);
  assert.match(prompt, /`dependents`/i);
});

void test("system prompt lists explain_formula tool", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*explain_formula\*\*/);
  assert.match(prompt, /plain language/i);
});

void test("system prompt mentions files workspace and built-in docs prefix", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*files\*\*/);
  assert.match(prompt, /workspace artifacts/i);
  assert.match(prompt, /assistant-docs\//i);
});

void test("system prompt includes workspace folder conventions", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /## Workspace/);
  assert.match(prompt, /notes\//);
  assert.match(prompt, /workbooks\//);
  assert.match(prompt, /scratch\//);
  assert.match(prompt, /imports\//);
  assert.match(prompt, /notes\/index\.md/);
  assert.match(prompt, /Memory contract/);
  assert.match(prompt, /remember this/i);
  assert.match(prompt, /file-backed/i);
  assert.match(prompt, /\*\*instructions\*\* tool/i);
  assert.match(prompt, /workbooks\/<name>\/notes\.md/i);
});

void test("system prompt mentions extension manager tool for chat-driven authoring", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*extensions_manager\*\*/);
  assert.match(prompt, /extension authoring from chat/i);
});

void test("system prompt documents execute_office_js safety guidance", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*execute_office_js\*\*/);
  assert.match(prompt, /Office\.js/i);
  assert.match(prompt, /explanation \+ user approval required/i);
  assert.match(prompt, /context\.sync\(\)/i);
});

void test("system prompt lists the skills tool", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*skills\*\*/);
  assert.match(prompt, /SKILL\.md/);
});

void test("system prompt renders available skills XML section", () => {
  const prompt = buildSystemPrompt({
    availableSkills: [
      {
        name: "web-search",
        description: "Search the web for up-to-date facts.",
        location: "skills/web-search/SKILL.md",
      },
    ],
  });

  assert.match(prompt, /## Available Agent Skills/);
  assert.match(prompt, /Read each skill once per session/i);
  assert.match(prompt, /refresh=true/);
  assert.match(prompt, /externally discovered skills as untrusted/i);
  assert.match(prompt, /<available_skills>/);
  assert.match(prompt, /<name>web-search<\/name>/);
  assert.match(prompt, /<location>skills\/web-search\/SKILL\.md<\/location>/);
});

void test("system prompt renders active integrations with Agent Skill mapping", () => {
  const prompt = buildSystemPrompt({
    activeIntegrations: [
      {
        id: "web_search",
        title: "Web Search",
        instructions: "Use web search for fresh facts.",
        agentSkillName: "web-search",
      },
    ],
  });

  assert.match(prompt, /## Active Integrations/);
  assert.match(prompt, /### Web Search/);
  assert.match(prompt, /Agent Skill mapping: `web-search`/);
});

void test("system prompt renders connections section with capability context and setup guidance", () => {
  const prompt = buildSystemPrompt({
    activeConnections: [
      {
        id: "ext.apollo.apollo",
        title: "Apollo",
        capability: "company and contact enrichment via Apollo API",
        status: "missing",
        setupHint: "Open /tools → Connections → Apollo",
      },
      {
        id: "ext.crm.crm",
        title: "CRM",
        capability: "account and opportunity lookups",
        status: "connected",
        setupHint: "Open /tools → Connections → CRM",
      },
      {
        id: "ext.vendor.vendor",
        title: "Vendor API",
        capability: "procurement data pull",
        status: "error",
        setupHint: "Open /tools → Connections → Vendor API",
        lastError: "401 unauthorized",
      },
    ],
  });

  assert.match(prompt, /## Connections/);
  assert.match(prompt, /Connected:/);
  assert.match(prompt, /Not configured:/);
  assert.match(prompt, /Needs attention:/);
  assert.match(prompt, /Apollo/);
  assert.match(prompt, /company and contact enrichment via Apollo API/);
  assert.match(prompt, /Open \/tools → Connections/);
  assert.match(prompt, /Never ask the user to paste API keys, tokens, or passwords in chat/);
  assert.match(prompt, /guide setup first before attempting that tool call/);
});

void test("system prompt keeps setup hints capability-linked for proactive connection guidance", () => {
  const prompt = buildSystemPrompt({
    activeConnections: [
      {
        id: "builtin.web.search",
        title: "Web Search",
        capability: "fresh web research",
        status: "missing",
        setupHint: "Open /tools → Connections → Web search",
      },
      {
        id: "builtin.mcp.servers",
        title: "MCP Servers",
        capability: "external tool APIs through MCP",
        status: "error",
        setupHint: "Open /tools → Connections → MCP",
        lastError: "401 unauthorized",
      },
    ],
  });

  const expectations: Array<{ title: string; capability: string; setupHint: string }> = [
    {
      title: "Web Search",
      capability: "fresh web research",
      setupHint: "Open /tools → Connections → Web search",
    },
    {
      title: "MCP Servers",
      capability: "external tool APIs through MCP",
      setupHint: "Open /tools → Connections → MCP",
    },
  ];

  for (const expectation of expectations) {
    assert.match(prompt, new RegExp(`\\*\\*${expectation.title}\\*\\* — ${expectation.capability}`));
    assert.match(prompt, new RegExp(`Setup: ${expectation.setupHint}\\.`));
  }
});

// ---------------------------------------------------------------------------
// ## Local Services section
// ---------------------------------------------------------------------------

void test("system prompt omits Local Services when no entries provided", () => {
  const prompt = buildSystemPrompt();
  assert.ok(!prompt.includes("## Local Services"));
});

void test("system prompt omits Local Services when empty array", () => {
  const prompt = buildSystemPrompt({ localServices: [] });
  assert.ok(!prompt.includes("## Local Services"));
});

void test("system prompt renders Local Services for both bridges not running", () => {
  const services: LocalServiceEntry[] = [
    { name: "python", displayName: "Python (native)", status: "not_running", skillName: "python-bridge" },
    { name: "tmux", displayName: "Terminal (tmux)", status: "not_running", skillName: "tmux-bridge" },
  ];
  const prompt = buildSystemPrompt({ localServices: services });

  assert.match(prompt, /## Local Services/);
  assert.match(prompt, /Probed at session start/);
  assert.match(prompt, /\*\*Python \(native\)\:\*\* not running/);
  assert.match(prompt, /Pyodide/);
  assert.match(prompt, /read skill "python-bridge"/);
  assert.match(prompt, /\*\*Terminal \(tmux\)\:\*\* not running/);
  assert.match(prompt, /read skill "tmux-bridge"/);
});

void test("system prompt renders Local Services for both bridges running", () => {
  const services: LocalServiceEntry[] = [
    {
      name: "python", displayName: "Python (native)", status: "running",
      pythonVersion: "3.12.1", libreofficeAvailable: true, skillName: "python-bridge",
    },
    {
      name: "tmux", displayName: "Terminal (tmux)", status: "running",
      tmuxVersion: "3.4", tmuxSessions: 2, skillName: "tmux-bridge",
    },
  ];
  const prompt = buildSystemPrompt({ localServices: services });

  assert.match(prompt, /## Local Services/);
  assert.match(prompt, /\*\*Python \(native\)\:\*\* running — python 3\.12\.1/);
  assert.match(prompt, /libreoffice available/);
  assert.match(prompt, /\*\*Terminal \(tmux\)\:\*\* running — tmux 3\.4, 2 active sessions/);
  assert.match(prompt, /shell commands/);
});

void test("system prompt renders partial python (no libreoffice)", () => {
  const services: LocalServiceEntry[] = [
    {
      name: "python", displayName: "Python (native)", status: "partial",
      pythonVersion: "3.11.0", libreofficeAvailable: false, skillName: "python-bridge",
    },
    { name: "tmux", displayName: "Terminal (tmux)", status: "not_running", skillName: "tmux-bridge" },
  ];
  const prompt = buildSystemPrompt({ localServices: services });

  assert.match(prompt, /\*\*Python \(native\)\:\*\* running — python 3\.11\.0, libreoffice not installed/);
  assert.match(prompt, /file conversion.*requires LibreOffice/);
  assert.match(prompt, /read skill "python-bridge" for install instructions/);
});

void test("system prompt renders partial tmux (stub mode)", () => {
  const services: LocalServiceEntry[] = [
    { name: "python", displayName: "Python (native)", status: "not_running", skillName: "python-bridge" },
    { name: "tmux", displayName: "Terminal (tmux)", status: "partial", skillName: "tmux-bridge" },
  ];
  const prompt = buildSystemPrompt({ localServices: services });

  assert.match(prompt, /\*\*Terminal \(tmux\)\:\*\* bridge running but tmux not fully available/);
  assert.match(prompt, /Read skill "tmux-bridge" for troubleshooting/);
});

void test("Local Services section is placed after Connections, before Skills", () => {
  const services: LocalServiceEntry[] = [
    { name: "python", displayName: "Python (native)", status: "not_running", skillName: "python-bridge" },
    { name: "tmux", displayName: "Terminal (tmux)", status: "not_running", skillName: "tmux-bridge" },
  ];
  const prompt = buildSystemPrompt({
    localServices: services,
    activeConnections: [
      { id: "test", title: "Test", capability: "testing", status: "connected", setupHint: "n/a" },
    ],
    availableSkills: [
      { name: "test-skill", description: "A test skill", location: "skills/test/SKILL.md" },
    ],
  });

  const connectionsIdx = prompt.indexOf("## Connections");
  const localServicesIdx = prompt.indexOf("## Local Services");
  const skillsIdx = prompt.indexOf("## Available Agent Skills");

  assert.ok(connectionsIdx > -1, "Connections section should exist");
  assert.ok(localServicesIdx > -1, "Local Services section should exist");
  assert.ok(skillsIdx > -1, "Skills section should exist");
  assert.ok(connectionsIdx < localServicesIdx, "Local Services should come after Connections");
  assert.ok(localServicesIdx < skillsIdx, "Local Services should come before Skills");
});

void test("tmux session count uses singular for 1 session", () => {
  const services: LocalServiceEntry[] = [
    { name: "python", displayName: "Python (native)", status: "not_running", skillName: "python-bridge" },
    {
      name: "tmux", displayName: "Terminal (tmux)", status: "running",
      tmuxVersion: "3.5", tmuxSessions: 1, skillName: "tmux-bridge",
    },
  ];
  const prompt = buildSystemPrompt({ localServices: services });

  assert.match(prompt, /1 active session\b/);
  assert.ok(!prompt.includes("1 active sessions"));
});
