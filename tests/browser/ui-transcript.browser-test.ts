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

interface RowFact { state: string | undefined; text: string }

async function readTurn(page: Page): Promise<{ copies: number; groups: number; title: string; rows: RowFact[]; thinking: string[] }> {
  return page.evaluate(() => {
    const host = document.querySelector('[data-gallery="transcript"]');
    if (!host) throw new Error("transcript section missing");
    const copies = [...host.querySelectorAll("assistant-message")]
      .filter((el) => el.textContent.includes("three blocks of the sheet")).length;
    const rows = [...host.querySelectorAll<HTMLElement>('.pi-turn-row[data-tool-name="read_range"]')].map((row) => ({
      state: row.dataset.state,
      text: row.querySelector(".pi-turn-row__sentence")?.textContent.replace(/\s+/g, " ").trim() ?? "",
    }));
    const thinking = [...host.querySelectorAll(".pi-thinking-label")].map((el) => el.textContent.trim());
    return {
      copies,
      groups: host.querySelectorAll("turn-group").length,
      title: host.querySelector(".pi-turn__title")?.textContent.trim() ?? "",
      rows,
      thinking,
    };
  });
}

void test("a streaming turn draws one turn group whose rows finish as each tool completes", async () => {
  const opened = await openBrowserPage(env, { path: "src/ui-gallery.html", viewport });
  const { page } = opened;
  try {
    await page.waitForFunction(() => Boolean(window.piTranscriptHarness));
    await page.evaluate(() => { void window.piTranscriptHarness?.start({ autoRelease: false }); });

    await page.waitForFunction(() => window.piTranscriptHarness?.agent.state.pendingToolCalls.size === 3);
    await page.waitForTimeout(200);
    let turn = await readTurn(page);
    assert.equal(turn.copies, 1, "the response with running tools is drawn once");
    assert.equal(turn.groups, 1, "one group for the turn");
    assert.equal(turn.title, "Working");
    assert.deepEqual(turn.rows.map((row) => row.state), ["live", "live", "live"]);
    assert.deepEqual(turn.rows.map((row) => row.text), ["Reading Sheet1!A1:B12", "Reading Sheet1!C1:D12", "Reading Sheet1!E1:F12"]);
    assert.match(turn.thinking[0] ?? "", /^Thought for \d+s$/, "reasoning ended before the tool calls");

    // The slowest tool is still running: the two finished rows read in the past tense now.
    await page.evaluate(() => {
      window.piTranscriptHarness?.release("read-2");
      window.piTranscriptHarness?.release("read-1");
    });
    await page.waitForFunction(() => window.piTranscriptHarness?.agent.state.pendingToolCalls.size === 1);
    await page.waitForTimeout(200);
    turn = await readTurn(page);
    assert.equal(turn.copies, 1);
    assert.deepEqual(turn.rows.map((row) => row.state), ["done", "done", "live"]);
    assert.equal(turn.rows[0]?.text, "Read Sheet1!A1:B12");

    await page.evaluate(() => window.piTranscriptHarness?.release("read-3"));
    await page.waitForFunction(() => window.piTranscriptHarness?.agent.state.isStreaming === false);
    await page.waitForTimeout(300);
    turn = await readTurn(page);
    assert.equal(turn.copies, 1);
    assert.equal(turn.groups, 1, "the second response joins the same turn");
    assert.match(turn.title, /^Worked for \d+s$/);
    assert.ok(turn.thinking.every((label) => /^Thought for \d+s$/.test(label)), turn.thinking.join(" | "));

    // The answer started, so the group collapsed itself; its header opens it again.
    assert.deepEqual(turn.rows, [], "collapsed group hides its rows");
    const groupHeader = page.locator('[data-gallery="transcript"] .pi-turn__header');
    await groupHeader.click();
    await page.waitForTimeout(100);
    assert.equal(await groupHeader.getAttribute("aria-expanded"), "true");
    assert.deepEqual((await readTurn(page)).rows.map((row) => row.state), ["done", "done", "done"]);

    // A row drills into the call's input and result (its reference is a link, so open it from the chevron).
    const firstRow = page.locator('[data-gallery="transcript"] .pi-turn-row__chevron').first();
    await firstRow.click();
    await page.locator('[data-gallery="transcript"] .pi-turn-row__detail').first().waitFor({ state: "visible" });
    assert.match(await page.locator('[data-gallery="transcript"] .pi-turn-row__detail').first().innerText(), /12 rows of monthly totals/);

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
    const toast = page.locator("#pi-toast .pi-toast.visible").last();
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
  await page.locator("turn-group").waitFor({ state: "attached" });
}

async function readDetails(page: Page): Promise<{ openBodies: number; groupExpanded: string | null; thinkingExpanded: string | null }> {
  return page.evaluate(() => ({
    openBodies: document.querySelectorAll(".pi-turn-row__detail").length,
    groupExpanded: document.querySelector(".pi-turn__header")?.getAttribute("aria-expanded") ?? null,
    thinkingExpanded: document.querySelector("thinking-block .thinking-header")?.getAttribute("aria-expanded") ?? null,
  }));
}

void test("Ctrl+O switches details mode for every row, survives re-renders and reloads", async () => {
  const opened = await openTaskpane(env, { viewport, clientId: "ui-transcript-details" });
  const { page } = opened;
  try {
    await seedRestoredTurn(page);
    await page.waitForTimeout(200);
    // A restored turn with no answer keeps its group open; its rows start closed.
    assert.deepEqual(await readDetails(page), { openBodies: 0, groupExpanded: "true", thinkingExpanded: "false" });

    // First press expands: details start collapsed, matching the rows.
    await page.keyboard.press("Control+o");
    await page.waitForTimeout(200);
    assert.deepEqual(await readDetails(page), { openBodies: 3, groupExpanded: "true", thinkingExpanded: "true" });
    assert.match(await page.locator("#pi-toast").innerText(), /Ctrl\+O|⌘O/);

    // A re-render (tab switch, response moving into the list) keeps the mode.
    await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      sidebar.agent.state.messages = sidebar.agent.state.messages.slice();
      sidebar.syncFromAgent();
    });
    await page.waitForTimeout(200);
    assert.equal((await readDetails(page)).openBodies, 3);

    // A hand toggle overrides one row until the next Ctrl+O.
    await page.locator(".pi-turn-row__chevron").first().click();
    await page.waitForFunction(() => document.querySelectorAll(".pi-turn-row__detail").length === 2);
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
    assert.deepEqual(await readDetails(page), { openBodies: 0, groupExpanded: "true", thinkingExpanded: "false" });
  } finally {
    await opened.finish();
  }
});

/* ── Turn groups across tab switches ── */

async function seedTab(page: Page, tag: string, turns: string[][]): Promise<void> {
  await page.evaluate(({ tag, turns }) => {
    const sidebar = document.querySelector("pi-sidebar");
    const usage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } };
    for (const [turnIndex, names] of turns.entries()) {
      const timestamp = 1000 + turnIndex * 10;
      const calls = names.map((name, index) => ({ type: "toolCall", id: `${tag}-${turnIndex}-${index}`, name, arguments: { range: `Sheet1!A${index + 1}` } }));
      sidebar.agent.state.messages.push(
        { role: "user", content: `${tag} ask ${turnIndex}`, timestamp: timestamp - 1 },
        { role: "assistant", content: [{ type: "text", text: `${tag} turn` }, ...calls], api: "openai-responses", provider: "openai", model: "fixture", usage, stopReason: "toolUse", timestamp },
        ...calls.map((call) => ({ role: "toolResult", toolCallId: call.id, toolName: call.name, content: [{ type: "text", text: "ok" }], isError: false, timestamp: timestamp + 1 })),
      );
    }
    sidebar.syncFromAgent();
  }, { tag, turns });
}

async function readGroups(page: Page): Promise<{ rows: number; groups: string[] }> {
  return page.evaluate(() => ({
    rows: document.querySelectorAll("pi-sidebar .pi-turn-row").length,
    groups: [...document.querySelectorAll("pi-sidebar turn-group")].map((group) =>
      `${group.querySelector(".pi-turn__title")?.textContent?.trim() ?? ""}:${group.querySelectorAll(".pi-turn-row").length}`),
  }));
}

void test("turn groups keep exactly their rows across tab switches", async () => {
  const opened = await openTaskpane(env, { viewport, clientId: "ui-transcript-tabs" });
  const { page } = opened;
  try {
    await seedTab(page, "A", [["read_range", "read_range", "read_range"], ["write_cells", "write_cells"]]);
    await page.locator(".pi-session-tabs__new").click();
    await page.waitForTimeout(300);
    await seedTab(page, "B", [["read_range", "read_range", "read_range", "read_range"], ["read_range"]]);

    const tabs = page.locator(".pi-session-tab__main");
    const expected = {
      A: { rows: 5, groups: ["Worked for 0s:3", "Worked for 0s:2"] },
      B: { rows: 5, groups: ["Worked for 0s:4", "Worked for 0s:1"] },
    };
    for (let switchIndex = 0; switchIndex < 6; switchIndex++) {
      const tab = switchIndex % 2 === 0 ? "A" : "B";
      await tabs.nth(switchIndex % 2).click();
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
      // One turn per call, so each row's sentence can be read on its own.
      for (const [index, item] of cases.entries()) {
        sidebar.agent.state.messages.push(
          { role: "user", content: `stop ${index}`, timestamp: 2000 + index * 10 },
          { role: "assistant", content: [{ type: "toolCall", id: item.id, name: item.name, arguments: item.args }], api: "openai-responses", provider: "openai", model: "fixture", usage, stopReason: "toolUse", timestamp: 2001 + index * 10 },
          { role: "toolResult", toolCallId: item.id, toolName: item.name, content: [{ type: "text", text: item.text }], details: item.details, isError: item.isError, timestamp: 2002 + index * 10 },
        );
      }
      sidebar.syncFromAgent();
    }, cases);
    await page.waitForTimeout(200);
    const rows = await page.locator(".pi-turn-row[data-tool-name]").evaluateAll((items) => items.map((item) => [
      item.querySelector(".pi-turn-row__sentence")?.textContent?.replace(/\s+/g, " ").trim() ?? "",
      item.querySelector(".pi-turn-row__meta")?.textContent?.trim() ?? "",
    ].filter(Boolean).join(" | ")));
    assert.deepEqual(rows, [
      "Stopped | 0 cells",
      "Stopped Stop!A1:B40 | 20 cells",
      "Stopped | 0 cells",
      "Stopped Stop!C1:C9 | 9 cells",
      "Edited Stop!A1:B2",
    ]);
  } finally {
    await opened.finish();
  }
});

void test("turn receipts render every ending and dispatch typed Review and Undo actions", async () => {
  const opened = await openBrowserPage(env, { path: "src/ui-gallery.html", viewport });
  const { page } = opened;
  try {
    const receipts = page.locator('[data-gallery="turn-changes"] .pi-turn-changes');
    await receipts.first().waitFor();
    assert.deepEqual(await receipts.evaluateAll((items) => items.map((item) => item.getAttribute("data-outcome"))), ["finished", "stopped", "failed"]);
    assert.match(await receipts.first().innerText(), /Changed 40 cells/);
    assert.match(await receipts.last().innerText(), /Some changes have no backup/);
    for (const label of ["Review", "Undo"]) {
      const action = page.evaluate(() => new Promise((resolve) => {
        document.addEventListener("pi:turn-change-action", (event) => {
          if (!(event instanceof CustomEvent)) return;
          resolve(event.detail);
        }, { once: true });
      }));
      await receipts.first().getByRole("button", { name: label, exact: true }).click();
      assert.deepEqual(await action, {
        action: label === "Review" ? "review" : "undo",
        summary: { role: "turnChanges", turnId: "gallery-turn-0", timestamp: 0, outcome: "finished", changes: [
          { address: "Sheet1!B2:B40", changedCount: 39, backupIds: ["gallery-backup-0"] },
          { address: "Summary!A1", changedCount: 1, backupIds: ["gallery-summary-0"] },
        ] },
      });
    }
  } finally {
    await opened.finish();
  }
});

void test("Office.js-only and mixed turns explicitly disclose the limit of turn undo", async () => {
  const opened = await openBrowserPage(env, { path: "src/ui-gallery.html", viewport });
  const { page } = opened;
  try {
    await page.locator('[data-gallery="turn-changes"] message-list').waitFor();
    await page.evaluate(async () => {
      const modulePath = "/src/messages/turn-changes.ts";
      const { summarizeTurnChanges } = await import(modulePath);
      const list = document.querySelector('[data-gallery="turn-changes"]')?.querySelector("message-list");
      if (!list) throw new Error("Turn receipt gallery missing");
      const shared = { timestamp: 1, outcome: "finished", ranOfficeJs: true, audit: [] };
      list.messages = [
        summarizeTurnChanges({ ...shared, turnId: "js-only", toolCallIds: new Set(["js"]), backups: [] }),
        summarizeTurnChanges({ ...shared, turnId: "mixed", toolCallIds: new Set(["core", "js"]),
          officeJsRanges: ["Sheet1!X12", "'Q3 Plan'!B2:C3", "Sheet1!A1", "Sheet1!A2", "Sheet1!A3", "Sheet1!A4", "Sheet1!A5"], backups: [
          { id: "core-backup", at: 0, toolCallId: "core", address: "Sheet1!D7", changedCount: 1 },
        ] }),
      ];
    });
    const receipts = page.locator('[data-gallery="turn-changes"] .pi-turn-changes');
    await page.waitForFunction(() => document.querySelectorAll('[data-gallery="turn-changes"] .pi-turn-changes').length === 2);
    for (const receipt of await receipts.all()) {
      assert.match(await receipt.innerText(), /not covered by Undo this turn\. Use \/backup/);
    }
    // No change events arrived for the code: the warning cannot name ranges.
    assert.match(await receipts.first().innerText(), /Ran Office\.js code; not covered/);
    assert.equal(await receipts.first().getByRole("button", { name: "Undo", exact: true }).count(), 0);
    assert.doesNotMatch(await receipts.first().innerText(), /Changed 0 cells/);
    await receipts.last().getByRole("button", { name: "Undo", exact: true }).click();
    assert.match(await receipts.last().innerText(), /Changed 1 cells/);
    assert.match(await receipts.last().innerText(),
      /Office\.js code changed Sheet1!X12, 'Q3 Plan'!B2:C3, Sheet1!A1, Sheet1!A2, Sheet1!A3, Sheet1!A4 \+1; not covered by Undo this turn/);
  } finally {
    await opened.finish();
  }
});
