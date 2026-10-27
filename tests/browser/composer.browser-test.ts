import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { enterInput, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

async function withTaskpane(run: (page: Page) => Promise<void>, viewportHeight = 720): Promise<void> {
  const opened = await openTaskpane(env, {
    viewport: { width: 420, height: viewportHeight },
    clientId: "composer-browser-client",
  });
  try {
    await run(opened.page);
  } finally {
    await opened.finish();
  }
}

/**
 * Gives the active tab enough history to compact and a model stream that never
 * answers, so `/compact` keeps the tab busy until it is stopped.
 */
async function holdCompaction(page: Page): Promise<void> {
  await page.evaluate(() => {
    const sidebar = document.querySelector("pi-sidebar");
    if (!sidebar || !("agent" in sidebar)) throw new Error("sidebar missing");
    // The sidebar's agent is the active tab's Agent; the test drives it directly.
    const agent = sidebar.agent as {
      state: { model: { contextWindow: number }; messages: Array<{ role: string; content: string; timestamp: number }> };
      getApiKey: () => Promise<string>;
      streamFunction: () => Promise<never>;
    };
    agent.state.messages = ["one", "two", "three", "four", "five"].map((text, index) => ({
      role: "user",
      content: `${text} `.repeat(100),
      timestamp: index + 1,
    }));
    // A tiny window leaves older messages to summarize.
    agent.state.model = { ...agent.state.model, contextWindow: 300 };
    agent.getApiKey = () => Promise.resolve("test-key");
    agent.streamFunction = () => new Promise<never>(() => {});
  });
}

void test("Alt+Up edits only the last queued prompt, a row's × removes it, and /compact keeps running", async () => {
  await withTaskpane(async (page) => {
    await holdCompaction(page);
    const first = "first queued action, long enough that the strip used to cut it off after forty-seven characters";

    await enterInput(page, "/compact");
    await page.locator("pi-working-indicator").getByText("Compacting context…").waitFor({ state: "visible" });
    await enterInput(page, first);
    await page.locator("#pi-queue-display").getByText(first, { exact: true }).waitFor({ state: "attached" });
    await enterInput(page, "second queued action");
    await page.locator("#pi-queue-display").getByText("second queued action", { exact: true }).waitFor({ state: "visible" });

    const input = page.locator("pi-input textarea");
    await input.press("Alt+ArrowUp");

    assert.deepEqual(
      {
        composer: await input.inputValue(),
        rowTexts: await page.locator("#pi-queue-display .pi-queue__text").allTextContents(),
        stillCompacting: await page.locator("pi-working-indicator").getByText("Compacting context…").isVisible(),
      },
      {
        composer: "second queued action",
        rowTexts: [first],
        stillCompacting: true,
      },
    );

    await page.locator("#pi-queue-display .pi-queue__remove").click();
    await page.locator("#pi-queue-display").waitFor({ state: "detached" });
    assert.equal(await input.inputValue(), "second queued action");
    assert.equal(await page.evaluate(() => document.activeElement?.tagName), "TEXTAREA");
  });
});

void test("Stop and Escape both stop compaction and keep the prompt held behind it as an unsent row", async () => {
  for (const trigger of ["button", "escape"] as const) {
    await withTaskpane(async (page) => {
      await holdCompaction(page);
      const input = page.locator("pi-input textarea");
      await enterInput(page, "/compact");
      await page.locator("pi-working-indicator").getByText("Compacting context…").waitFor({ state: "visible" });
      await enterInput(page, "held prompt");
      await page.locator("#pi-queue-display").getByText("held prompt", { exact: true }).waitFor({ state: "visible" });

      if (trigger === "button") {
        await page.locator(".pi-input-btn--abort").click();
      } else {
        await input.press("Escape");
      }
      await page.locator(".pi-input-btn--send").waitFor({ state: "visible" });

      assert.deepEqual(
        {
          composer: await input.inputValue(),
          labels: await page.locator("#pi-queue-display .pi-queue__label").allTextContents(),
          rowTexts: await page.locator("#pi-queue-display .pi-queue__text").allTextContents(),
          working: await page.locator("pi-working-indicator").innerText(),
        },
        { composer: "", labels: ["Not sent"], rowTexts: ["held prompt"], working: "" },
        trigger,
      );

      await page.locator("#pi-queue-display .pi-queue__edit").click();
      await page.locator("#pi-queue-display").waitFor({ state: "detached" });
      assert.equal(await input.inputValue(), "held prompt", trigger);
    });
  }
});

void test("the send button runs slash commands like Enter and an unknown command keeps the draft", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.fill("/compact");
    await page.locator(".pi-input-btn--send").click();
    await page.locator("#pi-toast").getByText("Too few messages to compact").waitFor({ state: "visible" });
    const userMessages = await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar)) return -1;
      const agent = sidebar.agent as { state: { messages: Array<{ role: string }> } };
      return agent.state.messages.filter((message) => message.role === "user").length;
    });
    assert.deepEqual({ composer: await input.inputValue(), userMessages }, { composer: "", userMessages: 0 });

    await input.fill("/nosuchcmd");
    await input.press("Enter");
    await page.locator("#pi-toast").getByText("No command /nosuchcmd", { exact: false }).waitFor({ state: "visible" });
    assert.equal(await input.inputValue(), "/nosuchcmd");
  });
});

void test("an Enter that commits an IME composition does not send", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await page.locator("pi-input").evaluate((composer) => {
      composer.addEventListener("pi-send", () => { document.body.dataset.composerSent = "yes"; });
    });
    await input.fill("你好");
    await input.evaluate((textarea) => {
      textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", isComposing: true, bubbles: true, cancelable: true }));
    });
    assert.equal(await page.locator("body").getAttribute("data-composer-sent"), null);
    assert.equal(await input.inputValue(), "你好");
  });
});

void test("each tab keeps its own draft, and a starter hint adds to it", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.fill("draft for chat one");
    await page.locator(".pi-session-tabs__new").click();
    await page.locator(".pi-session-tab").nth(1).waitFor({ state: "visible" });
    await page.waitForFunction(() => document.querySelector<HTMLTextAreaElement>("pi-input textarea")?.value === "");

    await input.fill("draft for chat two");
    await page.locator(".pi-session-tab").nth(0).locator(".pi-session-tab__main").click();
    await page.waitForFunction(() => document.querySelector<HTMLTextAreaElement>("pi-input textarea")?.value === "draft for chat one");

    await page.locator(".pi-empty__hint").first().click();
    const draft = await input.inputValue();
    assert.ok(draft.startsWith("draft for chat one\n\n"), draft);
    assert.ok(draft.length > "draft for chat one\n\n".length, draft);

    await page.locator(".pi-session-tab").nth(1).locator(".pi-session-tab__main").click();
    await page.waitForFunction(() => document.querySelector<HTMLTextAreaElement>("pi-input textarea")?.value === "draft for chat two");
  });
});

void test("undo in a tab's composer never replays typing from another tab", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    // The browser's own undo command, which Ctrl/Cmd+Z runs in the composer.
    const undo = () => page.evaluate(() => { document.execCommand("undo"); });
    await input.click();
    await page.keyboard.type("my draft");
    await page.locator(".pi-session-tabs__new").click();
    await page.locator(".pi-session-tab").nth(1).waitFor({ state: "visible" });
    await page.locator(".pi-session-tab").nth(0).locator(".pi-session-tab__main").click();
    await page.waitForFunction(() => document.querySelector<HTMLTextAreaElement>("pi-input textarea")?.value === "my draft");
    await input.click();
    await page.keyboard.press("End");
    await page.keyboard.type("abc");
    await input.selectText();
    await page.keyboard.press("Backspace");

    await page.locator(".pi-session-tab").nth(1).locator(".pi-session-tab__main").click();
    await page.waitForFunction(() => document.querySelector<HTMLTextAreaElement>("pi-input textarea")?.value === "");
    await input.click();
    await undo();
    await undo();
    assert.equal(await input.inputValue(), "");

    // Undo still works for this tab's own typing.
    await page.keyboard.type("tab two");
    await undo();
    assert.equal(await input.inputValue(), "");
  });
});

void test("the composer trims a prompt and does not submit whitespace", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await page.locator("pi-input").evaluate((composer) => {
      composer.addEventListener("pi-send", (event) => {
        // The component owns this event contract; the assertion only exposes its browser payload for observation.
        const sendEvent = event as CustomEvent<{ text?: string }>;
        if (typeof sendEvent.detail?.text === "string") {
          document.body.dataset.lastComposerSend = sendEvent.detail.text;
        }
      });
    });
    await input.fill("   \n\t  ");
    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-last-composer-send"), null);

    await input.fill("  build a TSLA P&L  ");
    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-last-composer-send"), "build a TSLA P&L");
  });
});

void test("Shift+Enter edits while Enter submits a real prompt", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await page.locator("pi-input").evaluate((composer) => {
      composer.addEventListener("pi-send", () => { document.body.dataset.composerSent = "yes"; });
    });
    await input.fill("write A1");
    await input.press("Shift+Enter");
    assert.equal(await page.locator("body").getAttribute("data-composer-sent"), null);
    assert.equal(await input.inputValue(), "write A1\n");

    await input.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-composer-sent"), "yes");
  });
});

void test("composer auto-grow honors its CSS cap", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.fill(Array.from({ length: 50 }, () => "row").join("\n"));
    assert.equal(await input.evaluate((element) => element.style.height), "288px");
  }, 720);
});

void test("composer auto-grow uses the smaller short-pane cap", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.evaluate((element) => { element.style.maxHeight = "none"; });
    await input.fill(Array.from({ length: 50 }, () => "row").join("\n"));
    assert.equal(await input.evaluate((element) => element.style.height), "140px");
  }, 500);
});

void test("each composer edit resizes once and caret keys do not resize it", async () => {
  await withTaskpane(async (page) => {
    const input = page.locator("pi-input textarea");
    await input.click();
    await page.locator("pi-input").evaluate((composer) => {
      const textarea = composer.querySelector("textarea");
      if (!textarea) throw new Error("composer textarea missing");
      const counters = document.body.dataset;
      counters.composerResizes = "0";
      counters.composerEdits = "0";
      new MutationObserver((records) => {
        counters.composerResizes = String(Number(counters.composerResizes) + records.length);
      }).observe(textarea, { attributes: true, attributeFilter: ["style"] });
      composer.addEventListener("input", (event) => {
        if (event.target === composer) counters.composerEdits = String(Number(counters.composerEdits) + 1);
      });
    });
    const takeWork = () => page.evaluate(() => {
      const counters = document.body.dataset;
      const work = { resizes: Number(counters.composerResizes), edits: Number(counters.composerEdits) };
      counters.composerResizes = "0";
      counters.composerEdits = "0";
      return work;
    });

    // Each resize writes the height twice: "auto" to measure, then the measured height.
    await input.pressSequentially("abc");
    assert.deepEqual(await takeWork(), { resizes: 6, edits: 3 });

    for (const key of ["ArrowLeft", "ArrowLeft", "ArrowRight", "Shift", "Home"]) await input.press(key);
    assert.deepEqual(await takeWork(), { resizes: 0, edits: 0 });
    assert.equal(await input.inputValue(), "abc");
  });
});

void test("a growing composer keeps the transcript following its newest message", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar) || !("requestUpdate" in sidebar)) throw new Error("sidebar missing");
      // The sidebar's agent is the active tab's Agent; the test seeds its transcript.
      const agent = sidebar.agent as { state: { messages: Array<{ role: string; content: string; timestamp: number }> } };
      for (let index = 0; index < 40; index += 1) {
        agent.state.messages.push({ role: "user", content: `message ${index}`, timestamp: index + 1 });
      }
      (sidebar.requestUpdate as () => void)();
    });
    const distanceFromBottom = () => page.locator(".pi-messages").evaluate(
      (element) => element.scrollHeight - element.scrollTop - element.clientHeight,
    );
    await page.waitForFunction(() => {
      const element = document.querySelector(".pi-messages");
      return element !== null && element.scrollHeight > element.clientHeight * 2;
    });
    await page.locator(".pi-messages").evaluate((element) => { element.scrollTop = element.scrollHeight; });

    const input = page.locator("pi-input textarea");
    await input.click();
    for (let line = 0; line < 5; line += 1) {
      await page.keyboard.type(`g${line}`);
      await page.keyboard.press("Shift+Enter");
    }
    await page.waitForTimeout(200);
    assert.ok((await distanceFromBottom()) < 20, `transcript ${await distanceFromBottom()}px from the bottom`);

    // Still following: a new message scrolls into view.
    await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar) || !("requestUpdate" in sidebar)) throw new Error("sidebar missing");
      const agent = sidebar.agent as { state: { messages: Array<{ role: string; content: string; timestamp: number }> } };
      agent.state.messages.push({ role: "user", content: "newest\n\nmessage\n\nwith\n\nseveral\n\nlines", timestamp: 100 });
      (sidebar.requestUpdate as () => void)();
    });
    await page.waitForTimeout(300);
    assert.ok((await distanceFromBottom()) < 20, `after a new message, transcript ${await distanceFromBottom()}px from the bottom`);
  });
});
