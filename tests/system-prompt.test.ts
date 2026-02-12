import assert from "node:assert/strict";
import { test } from "node:test";

import { buildSystemPrompt } from "../src/prompt/system-prompt.ts";
import { resolveConventions } from "../src/conventions/store.ts";

void test("system prompt includes default placeholders when instructions are absent", () => {
  const prompt = buildSystemPrompt();

  assert.match(prompt, /## Persistent Instructions/);
  assert.match(prompt, /\(No user instructions set\.\)/);
  assert.match(prompt, /\(No workbook instructions set\.\)/);
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
    currencySymbol: "£",
    negativeStyle: "minus",
  });
  const prompt = buildSystemPrompt({ conventions });

  assert.match(prompt, /Active convention overrides/);
  assert.match(prompt, /Currency: £/);
  assert.match(prompt, /Negatives: minus sign/);
});

void test("system prompt lists the conventions tool", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*conventions\*\*/);
});

void test("system prompt mentions optional files workspace capability", () => {
  const prompt = buildSystemPrompt();
  assert.match(prompt, /\*\*files\*\*/);
  assert.match(prompt, /workspace artifacts/i);
});
