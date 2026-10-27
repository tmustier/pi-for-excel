/* oxlint-disable typescript/no-unsafe-call, typescript/no-unsafe-member-access, typescript/no-unsafe-return, typescript/no-unsafe-assignment -- Playwright evaluates first-party custom elements and the gallery harness, whose browser runtime types are unavailable to the Node test checker. */
import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { openBrowserPage, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

const viewport = { width: 420, height: 900 };

/* ── Live turn: real <pi-sidebar> + Agent, scripted stream (src/ui-gallery-transcript.ts) ── */

interface CardFact { where: "list" | "stream"; state: string | undefined; visible: boolean }

async function readTurn(page: Page): Promise<{ copies: number; cards: CardFact[]; thinking: string[] }> {
  return page.evaluate(() => {
    const host = document.querySelector('[data-gallery="transcript"]');
    if (!host) throw new Error("transcript section missing");
    const copies = [...host.querySelectorAll("assistant-message")]
      .filter((el) => el.textContent.includes("three blocks of the sheet")).length;
    const cards = [...host.querySelectorAll<HTMLElement>(".pi-tool-card[data-tool-name=read_range]")].map((card) => ({
      where: card.closest("streaming-message-container") ? "stream" as const : "list" as const,
      state: card.dataset.state,
      visible: card.getBoundingClientRect().height > 0,
    }));
    const thinking = [...host.querySelectorAll(".pi-thinking-label")].map((el) => el.textContent.trim());
    return { copies, cards, thinking };
  });
}

void test("a streaming turn draws each response once and finishes each tool card as it completes", async () => {
  const opened = await openBrowserPage(env, { path: "src/ui-gallery.html", viewport });
  const { page } = opened;
  try {
    await page.waitForFunction(() => Boolean(window.piTranscriptHarness));
    await page.evaluate(() => { void window.piTranscriptHarness?.start({ autoRelease: false }); });

    await page.waitForFunction(() => window.piTranscriptHarness?.agent.state.pendingToolCalls.size === 3);
    await page.waitForTimeout(200);
    let turn = await readTurn(page);
    assert.equal(turn.copies, 1, "the response with running tools is drawn once");
    assert.deepEqual(turn.cards.map((card) => card.state), ["inprogress", "inprogress", "inprogress"]);
    assert.match(turn.thinking[0] ?? "", /^Thought for \d+s$/, "reasoning ended before the tool calls");

    // The slowest tool is still running: the two finished cards stop spinning now.
    await page.evaluate(() => {
      window.piTranscriptHarness?.release("read-2");
      window.piTranscriptHarness?.release("read-1");
    });
    await page.waitForFunction(() => window.piTranscriptHarness?.agent.state.pendingToolCalls.size === 1);
    await page.waitForTimeout(200);
    turn = await readTurn(page);
    assert.equal(turn.copies, 1);
    assert.deepEqual(turn.cards.map((card) => card.state), ["complete", "complete", "inprogress"]);

    await page.evaluate(() => window.piTranscriptHarness?.release("read-3"));
    await page.waitForFunction(() => window.piTranscriptHarness?.agent.state.isStreaming === false);
    await page.waitForTimeout(300);
    turn = await readTurn(page);
    assert.equal(turn.copies, 1);
    assert.deepEqual(turn.cards.map((card) => card.state), ["complete", "complete", "complete"]);
    assert.ok(turn.thinking.every((label) => /^Thought for \d+s$/.test(label)), turn.thinking.join(" | "));

    // Three reads group and start collapsed: the group really hides its cards.
    assert.deepEqual(turn.cards.map((card) => card.visible), [false, false, false]);
    const groupHeader = page.locator('[data-gallery="transcript"] .pi-tool-group__header');
    await groupHeader.click();
    await page.waitForTimeout(100);
    assert.equal(await groupHeader.getAttribute("aria-expanded"), "true");
    assert.deepEqual((await readTurn(page)).cards.map((card) => card.visible), [true, true, true]);

    // Thinking header is a keyboard control.
    const thinkingHeader = page.locator('[data-gallery="transcript"] thinking-block .thinking-header').first();
    await thinkingHeader.focus();
    await page.keyboard.press("Enter");
    assert.equal(await thinkingHeader.getAttribute("aria-expanded"), "true");
    await page.locator('[data-gallery="transcript"] thinking-block markdown-block').first().waitFor({ state: "visible" });

    // Fence labels with symbols keep their code bar; single newlines stay line breaks.
    const languages = await page.locator('[data-gallery="transcript"] message-list .pi-code__lang').allTextContents();
    assert.deepEqual(languages, ["c++", "f#"]);
    assert.equal(await page.locator('[data-gallery="transcript"] user-message .pi-markdown').innerText(), "Summarise Sheet1.\nKeep it short.");

    // No Excel here: a cell link fails visibly, without an unhandled rejection (finish() checks page errors).
    await page.locator('[data-gallery="transcript"] .pi-cell-ref:visible').first().click();
    const toast = page.locator("#pi-toast.visible");
    await toast.waitFor({ state: "visible" });
    assert.match(await toast.innerText(), /Sheet1!A1:B12/);
  } finally {
    await opened.finish();
  }
});

/* ── Details mode (Ctrl/Cmd+O) in the real taskpane ── */

async function seedRestoredTurn(page: Page): Promise<void> {
  await page.evaluate(() => {
    const sidebar = document.querySelector("pi-sidebar");
    const usage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } };
    const ranges = ["Sheet1!A1:B2", "Sheet1!C1:D2", "Sheet1!E1:F2"];
    sidebar.agent.state.messages.push(
      {
        role: "assistant",
        content: [
          { type: "thinking", thinking: "Read the three blocks." },
          ...ranges.map((range, index) => ({ type: "toolCall", id: `details-${index}`, name: "read_range", arguments: { range } })),
        ],
        api: "openai-responses", provider: "openai", model: "fixture", usage, stopReason: "toolUse", timestamp: 1000,
      },
      ...ranges.map((range, index) => ({
        role: "toolResult", toolCallId: `details-${index}`, toolName: "read_range",
        content: [{ type: "text", text: `Read ${range}` }], isError: false, timestamp: 1001 + index,
      })),
    );
    sidebar.syncFromAgent();
  });
  await page.locator(".pi-tool-group").waitFor({ state: "attached" });
}

async function readDetails(page: Page): Promise<{ openBodies: number; groupExpanded: string | null; thinkingExpanded: string | null }> {
  return page.evaluate(() => ({
    openBodies: [...document.querySelectorAll<HTMLElement>(".pi-tool-card__body")]
      .filter((body) => body.getBoundingClientRect().height > 0).length,
    groupExpanded: document.querySelector(".pi-tool-group__header")?.getAttribute("aria-expanded") ?? null,
    thinkingExpanded: document.querySelector("thinking-block .thinking-header")?.getAttribute("aria-expanded") ?? null,
  }));
}

void test("Ctrl+O switches details mode for every card, survives re-renders and reloads", async () => {
  const opened = await openTaskpane(env, { viewport, clientId: "ui-transcript-details" });
  const { page } = opened;
  try {
    await seedRestoredTurn(page);
    await page.waitForTimeout(200);
    assert.deepEqual(await readDetails(page), { openBodies: 0, groupExpanded: "false", thinkingExpanded: "false" });

    // First press expands: details start collapsed, matching the cards.
    await page.keyboard.press("Control+o");
    await page.waitForTimeout(200);
    assert.deepEqual(await readDetails(page), { openBodies: 3, groupExpanded: "true", thinkingExpanded: "true" });
    assert.match(await page.locator("#pi-toast").innerText(), /Ctrl\+O/);

    // A re-render (tab switch, response moving into the list) keeps the mode.
    await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      sidebar.agent.state.messages = sidebar.agent.state.messages.slice();
      sidebar.syncFromAgent();
    });
    await page.waitForTimeout(200);
    assert.equal((await readDetails(page)).openBodies, 3);

    // A hand toggle overrides one card until the next Ctrl+O.
    await page.locator(".pi-tool-card__toggle").first().click();
    // The body animates closed; wait for the transition rather than a fixed delay.
    await page.waitForFunction(() => document.querySelectorAll(".pi-tool-card__body--collapsed").length === 1);
    await page.waitForTimeout(400);
    assert.equal((await readDetails(page)).openBodies, 2);

    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    const welcome = page.locator("#pi-welcome-login-overlay");
    await welcome.waitFor({ state: "visible", timeout: 10_000 });
    await welcome.dispatchEvent("click");
    await welcome.waitFor({ state: "detached" });
    await seedRestoredTurn(page);
    await page.waitForTimeout(200);
    assert.deepEqual(await readDetails(page), { openBodies: 3, groupExpanded: "true", thinkingExpanded: "true" }, "mode is saved");

    await page.keyboard.press("Control+o");
    await page.waitForTimeout(200);
    assert.deepEqual(await readDetails(page), { openBodies: 0, groupExpanded: "false", thinkingExpanded: "false" });
  } finally {
    await opened.finish();
  }
});

/* ── Tool groups across tab switches ── */

async function seedTab(page: Page, tag: string, turns: string[][]): Promise<void> {
  await page.evaluate(({ tag, turns }) => {
    const sidebar = document.querySelector("pi-sidebar");
    const usage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } };
    for (const [turnIndex, names] of turns.entries()) {
      const timestamp = 1000 + turnIndex * 10;
      const calls = names.map((name, index) => ({ type: "toolCall", id: `${tag}-${turnIndex}-${index}`, name, arguments: { range: `Sheet1!A${index + 1}` } }));
      sidebar.agent.state.messages.push(
        { role: "assistant", content: [{ type: "text", text: `${tag} turn` }, ...calls], api: "openai-responses", provider: "openai", model: "fixture", usage, stopReason: "toolUse", timestamp },
        ...calls.map((call) => ({ role: "toolResult", toolCallId: call.id, toolName: call.name, content: [{ type: "text", text: "ok" }], isError: false, timestamp: timestamp + 1 })),
      );
    }
    sidebar.syncFromAgent();
  }, { tag, turns });
}

async function readGroups(page: Page): Promise<{ cards: number; groups: string[] }> {
  return page.evaluate(() => ({
    cards: document.querySelectorAll("pi-sidebar tool-message").length,
    groups: [...document.querySelectorAll(".pi-tool-group")].map((group) =>
      `${group.querySelector(".pi-tool-group__label")?.textContent ?? ""}:${group.querySelectorAll("tool-message").length}`),
  }));
}

void test("tool groups keep exactly their cards across tab switches", async () => {
  const opened = await openTaskpane(env, { viewport, clientId: "ui-transcript-tabs" });
  const { page } = opened;
  try {
    await seedTab(page, "A", [["read_range", "read_range", "read_range"], ["write_cells", "write_cells"]]);
    await page.locator(".pi-session-tabs__new").click();
    await page.waitForTimeout(300);
    await seedTab(page, "B", [["read_range", "read_range", "read_range", "read_range"], ["read_range"]]);

    const tabs = page.locator(".pi-session-tab__main");
    const expected = {
      A: { cards: 5, groups: ["3 reads:3", "2 edits:2"] },
      B: { cards: 5, groups: ["4 reads:4"] },
    };
    for (let switchIndex = 0; switchIndex < 6; switchIndex++) {
      const tab = switchIndex % 2 === 0 ? "A" : "B";
      await tabs.nth(switchIndex % 2).click();
      await page.locator(".pi-tool-group__header").first().click();
      await page.waitForTimeout(150);
      assert.deepEqual(await readGroups(page), expected[tab], `switch ${switchIndex + 1} to ${tab}`);
    }
  } finally {
    await opened.finish();
  }
});

/* ── Stopped writes ── */

void test("a stopped write says how far it got instead of claiming success", async () => {
  const opened = await openTaskpane(env, { viewport, clientId: "ui-transcript-stopped" });
  const { page } = opened;
  try {
    // Details as write_cells / fill_formula (#791) and python_transform_range (#797) report a Stop.
    const cases = [
      { id: "s-1", name: "write_cells", args: { start_cell: "Stop!A1", values: [[1]] }, isError: true,
        text: "⏹ **Stopped before writing** — nothing changed in Stop!A1:B2.",
        details: { kind: "write_cells", blocked: false, address: "Stop!A1:B2", stopped: true, writtenCount: 0 } },
      { id: "s-2", name: "write_cells", args: { start_cell: "Stop!A1", values: [[1]] }, isError: true,
        text: "⏹ **Stopped after writing Stop!A1:B10** (rows 1–10 of Stop!A1:B40) — the change stands.\nThe remaining 30 row(s) were not written.",
        details: { kind: "write_cells", blocked: false, address: "Stop!A1:B40", stopped: true, writtenCount: 20 } },
      { id: "s-3", name: "python_transform_range", args: { range: "Stop!A1:A3", code: "x" }, isError: true,
        text: "⏹ **Stopped before writing** — nothing changed in Stop!D1.",
        details: { kind: "python_transform_range", blocked: false, outputAddress: "Stop!D1", stopped: true } },
      { id: "s-4", name: "fill_formula", args: { range: "Stop!C1:C9", formula: "=A1" }, isError: false,
        text: "⏹ **Stopped after filling Stop!C1:C9** — the change stands. It was not verified because Stop was pressed; read the range before relying on it.",
        details: { kind: "fill_formula", blocked: false, address: "Stop!C1:C9", verification: "failed", stopped: true, writtenCount: 9 } },
      { id: "s-5", name: "write_cells", args: { start_cell: "Stop!A1", values: [[1]] }, isError: false,
        text: "Written to **Stop!A1:B2**",
        details: { kind: "write_cells", blocked: false, address: "Stop!A1:B2" } },
    ];
    await page.evaluate((cases) => {
      const sidebar = document.querySelector("pi-sidebar");
      const usage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } };
      // One response per call, so finished calls of the same tool do not group.
      for (const [index, item] of cases.entries()) {
        sidebar.agent.state.messages.push(
          { role: "assistant", content: [{ type: "toolCall", id: item.id, name: item.name, arguments: item.args }], api: "openai-responses", provider: "openai", model: "fixture", usage, stopReason: "toolUse", timestamp: 2000 + index * 10 },
          { role: "toolResult", toolCallId: item.id, toolName: item.name, content: [{ type: "text", text: item.text }], details: item.details, isError: item.isError, timestamp: 2001 + index * 10 },
        );
      }
      sidebar.syncFromAgent();
    }, cases);
    await page.waitForTimeout(200);
    const titles = await page.locator(".pi-tool-card__title").allInnerTexts();
    assert.deepEqual(titles.map((title) => title.replace(/\s+/g, " ").trim()), [
      "Stopped — nothing written",
      "Stopped Stop!A1:B40 — 20 written",
      "Stopped — nothing written",
      "Stopped Stop!C1:C9 — 9 written",
      "Edit Stop!A1:B2",
    ]);
  } finally {
    await opened.finish();
  }
});
