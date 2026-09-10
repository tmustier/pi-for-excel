/* eslint-disable @typescript-eslint/no-unsafe-call, @typescript-eslint/no-unsafe-member-access -- Playwright evaluates first-party custom elements whose browser runtime types are unavailable to the Node test checker. */
import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, { viewport: { width: 420, height: 720 }, clientId: "ui-markdown-browser" });
  try { await run(opened.page); } finally { await opened.finish(); }
}

const rows = [
  { name: "thematic break", markdown: "---\nSection intro\n---\n\n## Thematic marker\nBody.", visible: "Section intro", hidden: null },
  { name: "YAML list", markdown: "---\ntags:\n  - excel\n  - formulas\n---\n# List marker\nBody.", visible: "List marker", hidden: "excel" },
  { name: "non-YAML block", markdown: "---\n## Intro: details\n---\n# Non-YAML marker", visible: "Intro: details", hidden: null },
  { name: "prose label", markdown: "---\nNote: this section is prose, not metadata\n---\n# Prose marker", visible: "Note: this section is prose", hidden: null },
  { name: "title-cased keys", markdown: "---\nTitle: Spreadsheet Skill\nDate: 2026-02-14\n---\n# Title marker", visible: "Title marker", hidden: "Spreadsheet Skill" },
  { name: "block scalar", markdown: "---\nname: skill\ndescription: >-\n  Multi-line description.\nmetadata:\n  integration-id: sample\n---\n# Scalar marker", visible: "Scalar marker", hidden: "Multi-line description" },
  { name: "UTF-8 BOM", markdown: "\uFEFF---\nname: skill\ndescription: Test\n---\n# BOM marker", visible: "BOM marker", hidden: "description: Test" },
] as const;

void test("assistant tool Markdown renders all seven frontmatter forms correctly", async () => {
  for (const [index, row] of rows.entries()) {
    await withTaskpane(async (page) => {
      await page.evaluate(({ index, markdown }) => {
        const sidebar = document.querySelector("pi-sidebar");
        sidebar.agent.state.messages.push(
          { role: "assistant", content: [{ type: "toolCall", id: `markdown-${index}`, name: "skills", arguments: { action: "read", name: `fixture-${index}` } }], api: "openai-responses", provider: "openai", model: "fixture", usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } }, stopReason: "toolUse", timestamp: index * 2 + 1 },
          { role: "toolResult", toolCallId: `markdown-${index}`, toolName: "skills", content: [{ type: "text", text: markdown }], isError: false, timestamp: index * 2 + 2 },
        );
        sidebar.syncFromAgent();
      }, { index, markdown: row.markdown });
      const card = page.locator('.pi-tool-card[data-tool-name="skills"]');
      await card.locator(".pi-tool-card__header").click();
      const rendered = card.locator(".pi-tool-card__markdown");
      await rendered.getByText(row.visible, { exact: false }).first().waitFor({ state: "visible" });
      const text = await rendered.innerText();
      assert.equal(text.includes(row.visible), true, row.name);
      if (row.hidden !== null) assert.equal(text.includes(row.hidden), false, row.name);
    });
  }
});
