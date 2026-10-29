/* oxlint-disable typescript/no-unsafe-call, typescript/no-unsafe-member-access -- Playwright evaluates first-party custom elements whose browser runtime types are unavailable to the Node test checker. */
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

void test("answer and table copy use TSV and expose manual fallback without claiming success", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(() => {
      Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: () => Promise.reject(new Error("Frame denied")) } });
      document.execCommand = () => false;
      const sidebar = document.querySelector("pi-sidebar");
      sidebar.agent.state.messages.push({
        role: "assistant", content: [
          { type: "thinking", thinking: "Private reasoning" },
          { type: "text", text: "**Forecast**\n\n| Region | Revenue |\n| --- | --- |\n| East | 17 |\n| West | 29 |\n\n```text\nraw <value>\n```" },
        ], api: "openai-responses", provider: "openai", model: "fixture",
        usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
        stopReason: "stop", timestamp: 1,
      });
      sidebar.syncFromAgent();
    });
    const dialog = page.locator(".pi-copy-text__value");
    await page.getByRole("button", { name: "Copy table", exact: true }).click();
    assert.equal(await dialog.inputValue(), "Region\tRevenue\nEast\t17\nWest\t29");
    await page.getByRole("button", { name: "Done", exact: true }).click();
    await page.getByRole("button", { name: "Copy answer", exact: true }).click();
    const answer = await dialog.inputValue();
    assert.match(answer, /Forecast/);
    assert.match(answer, /Region\tRevenue\nEast\t17\nWest\t29/);
    assert.match(answer, /raw <value>/);
    assert.doesNotMatch(answer, /Private reasoning|Copy table|Copy answer|```|\| Region/);
    await page.getByRole("button", { name: "Done", exact: true }).click();
    await page.locator("code-block").getByRole("button", { name: "Copy", exact: true }).click();
    assert.equal(await dialog.inputValue(), "raw <value>\n");
    assert.equal(await page.getByText("Copied!", { exact: true }).count(), 0);
  });
});

void test("only valid qualified references on known sheets become answer links, not code or ordinary links", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const context = {
        workbook: { worksheets: { items: [{ name: "Sheet1" }, { name: "Q3 Plan" }, { name: "Bob's" }], load: () => {} } },
        sync: () => Promise.resolve(),
      };
      Reflect.set(window, "Excel", { run: (callback: (ctx: typeof context) => Promise<Set<string>>) => callback(context) });
      const block = document.createElement("markdown-block");
      block.answer = true;
      block.content = "Sheet1!B2:B40, 'Q3 Plan'!$C$4:F9 and 'Bob''s'!A1. Unknown!A1 Sheet1!XFE1 Sheet1!A1048577 Sheet1!A0 B12 FY24. `Sheet1!C3` [Sheet1!D4](https://example.com)\n\n```text\nSheet1!E5\n```";
      document.querySelector(".pi-messages").append(block);
      await block.updateComplete;
    });
    const links = page.locator("markdown-block a.pi-cell-ref");
    await links.nth(2).waitFor();
    assert.deepEqual(await links.allTextContents(), ["Sheet1!B2:B40", "'Q3 Plan'!$C$4:F9", "'Bob''s'!A1"]);
    assert.equal(await page.getByRole("link", { name: "Sheet1!D4", exact: true }).getAttribute("href"), "https://example.com");
  });
});

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
      const skillRow = page.locator('.pi-turn-row[data-tool-name="skills"]');
      await skillRow.locator(".pi-turn-row__line").click();
      const rendered = skillRow.locator(".pi-tool-card__markdown");
      await rendered.getByText(row.visible, { exact: false }).first().waitFor({ state: "visible" });
      const text = await rendered.innerText();
      assert.equal(text.includes(row.visible), true, row.name);
      if (row.hidden !== null) assert.equal(text.includes(row.hidden), false, row.name);
    });
  }
});
