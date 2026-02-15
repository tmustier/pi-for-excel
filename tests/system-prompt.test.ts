import assert from "node:assert/strict";
import { test } from "node:test";

import { buildSystemPrompt } from "../src/prompt/system-prompt.ts";
import { resolveConventions } from "../src/conventions/store.ts";

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
  assert.match(prompt, /Current mode:\s*\*\*YOLO\*\*/);
  assert.match(prompt, /low-friction execution/i);
});

void test("system prompt renders Safe execution mode guidance", () => {
  const prompt = buildSystemPrompt({ executionMode: "safe" });

  assert.match(prompt, /Current mode:\s*\*\*Safe\*\*/);
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
