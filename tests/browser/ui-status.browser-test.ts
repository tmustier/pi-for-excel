/* oxlint-disable typescript/no-unsafe-assignment, typescript/no-unsafe-call, typescript/no-unsafe-member-access -- Playwright evaluates Vite-loaded application modules and custom elements whose browser runtime types are unavailable to the Node test checker. */
import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

async function withTaskpane(viewport: { width: number; height: number }, run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, { viewport, clientId: "ui-status-browser" });
  try { await run(opened.page); } finally { await opened.finish(); }
}

async function setContextUsage(page: Page, percent: number): Promise<void> {
  await page.evaluate((pct) => {
    const sidebar = document.querySelector("pi-sidebar");
    const agent = sidebar.agent;
    // A 200k window: yellow above 120k (75% of the 160k soft warning budget), red above 160k.
    agent.state.model = { ...agent.state.model, contextWindow: 200_000 };
    const tokens = pct * 2_000;
    agent.state.messages.push({
      role: "assistant",
      content: [{ type: "text", text: "usage fixture" }],
      api: "openai-responses",
      provider: "openai",
      model: "browser-fixture",
      usage: { input: tokens, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: tokens, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
      stopReason: "stop",
      timestamp: Date.now(),
    });
    document.dispatchEvent(new CustomEvent("pi:status-update"));
  }, percent);
  await page.locator(".pi-status-ctx__pct").filter({ hasText: `${percent}%` }).waitFor({ state: "visible" });
}

for (const row of [
  { percent: 60, warning: null, severity: null },
  { percent: 61, warning: "Context 61% full.", severity: "yellow" },
  { percent: 81, warning: "Context 81% full — responses may become less reliable.", severity: "red" },
  { percent: 101, warning: "Context is full — the next message will fail.", severity: "red" },
] as const) {
  void test(`context usage at ${row.percent}% shows the expected visible warning`, async () => {
    await withTaskpane({ width: 420, height: 720 }, async (page) => {
      await setContextUsage(page, row.percent);
      await page.locator(".pi-status-ctx--trigger").click();
      const warning = page.locator(".pi-status-popover__warning");
      if (row.warning === null) {
        assert.equal(await warning.count(), 0);
      } else {
        await warning.waitFor({ state: "visible" });
        assert.deepEqual({ text: await warning.innerText(), severityClass: await warning.getAttribute("class") }, { text: row.warning, severityClass: `pi-status-popover__warning pi-status-popover__warning--${row.severity}` });
      }
    });
  });
}

void test("context warning attributes normalize unknown severity to yellow in the visible popover", async () => {
  await withTaskpane({ width: 420, height: 720 }, async (page) => {
    await setContextUsage(page, 61);
    await page.locator(".pi-status-ctx--trigger").evaluate((trigger) => trigger.setAttribute("data-ctx-severity", "unknown"));
    await page.locator(".pi-status-ctx--trigger").click();
    assert.match(await page.locator(".pi-status-popover__warning").getAttribute("class") ?? "", /warning--yellow/);
  });
});

void test("Shift+Tab cycles thinking only from the composer text field", async () => {
  await withTaskpane({ width: 420, height: 720 }, async (page) => {
    const level = (): Promise<string | null> => page.locator(".pi-status-thinking").getAttribute("aria-label");
    const input = page.locator("pi-input textarea");
    await input.fill("draft");

    const before = await level();
    await input.press("Shift+Tab");
    const cycled = await level();
    assert.notEqual(cycled, before);
    assert.equal(await input.evaluate((el) => el === document.activeElement), true);

    // From Send, Shift+Tab moves focus back to the text field and leaves the level alone.
    await page.getByRole("button", { name: "Send" }).focus();
    await page.keyboard.press("Shift+Tab");
    assert.equal(await level(), cycled);
    assert.equal(await input.evaluate((el) => el === document.activeElement), true);
  });
});

void test("arrow keys inside a status popover do not switch tabs", async () => {
  await withTaskpane({ width: 420, height: 720 }, async (page) => {
    await page.getByRole("button", { name: "New tab" }).click();
    const active = page.locator(".pi-session-tab.is-active");
    await active.filter({ hasText: "Chat 2" }).waitFor();

    await page.locator(".pi-status-thinking").click();
    await page.keyboard.press("ArrowRight");
    await page.keyboard.press("ArrowLeft");

    assert.match(await active.textContent() ?? "", /Chat 2/);
    assert.equal(await page.locator(".pi-status-popover--thinking").isVisible(), true);
  });
});

void test("toasts sit above the composer dock as it grows", async () => {
  await withTaskpane({ width: 380, height: 622 }, async (page) => {
    const gap = await page.evaluate(async () => {
      // A queue strip is inserted into the dock above the input, as queue-display.ts does.
      const inputArea = document.querySelector(".pi-input-area");
      const strip = document.createElement("div");
      strip.style.height = "60px";
      inputArea.parentElement.insertBefore(strip, inputArea);
      await new Promise((resolve) => setTimeout(resolve, 200));
      const { showToast } = await import("/src/ui/toast.ts");
      showToast("Toast placement fixture");
      await new Promise((resolve) => setTimeout(resolve, 300));
      const toast = document.querySelector(".pi-toast-region").getBoundingClientRect();
      return strip.getBoundingClientRect().top - toast.bottom;
    });
    assert.ok(gap >= 0, `toast overlaps the dock by ${-gap}px`);
  });
});

void test("a toast with the same id replaces a long progress toast", async () => {
  await withTaskpane({ width: 380, height: 622 }, async (page) => {
    const messages = await page.evaluate(async () => {
      const { showToast } = await import("/src/ui/toast.ts");
      showToast("Unrelated toast");
      showToast("Compacting fixture…", { id: "compact", duration: 60_000 });
      showToast("Summarized fixture", { id: "compact" });
      return [...document.querySelectorAll("#pi-toast .pi-toast__message")].map((node) => node.textContent);
    });
    assert.deepEqual(messages, ["Unrelated toast", "Summarized fixture"]);
  });
});

void test("thinking popover stays visible inside a 290 by 400 taskpane", async () => {
  await withTaskpane({ width: 290, height: 400 }, async (page) => {
    await page.locator(".pi-status-thinking").click();
    const box = await page.locator(".pi-status-popover--thinking").boundingBox();
    assert.ok(box);
    assert.ok(box.x >= 8 && box.y >= 8 && box.x + box.width <= 282 && box.y + box.height <= 392);
  });
});

void test("thinking popover remains above its anchor when it fits", async () => {
  await withTaskpane({ width: 800, height: 900 }, async (page) => {
    const anchor = page.locator(".pi-status-thinking");
    await anchor.click();
    const [anchorBox, popoverBox] = await Promise.all([anchor.boundingBox(), page.locator(".pi-status-popover--thinking").boundingBox()]);
    assert.ok(anchorBox && popoverBox);
    assert.ok(popoverBox.y + popoverBox.height < anchorBox.y);
  });
});

void test("choosing a mode updates the label at once while the status bar is hovered and focused", async () => {
  await withTaskpane({ width: 380, height: 720 }, async (page) => {
    const mode = page.locator(".pi-status-mode");
    const initial = (await mode.innerText()).trim();
    const target = initial === "Auto" ? "Confirm" : "Auto";

    await mode.hover();
    await mode.focus();
    await page.keyboard.press("Enter");
    await page.getByRole("menuitemradio", { name: new RegExp(`^${target}`) }).waitFor({ state: "visible" });
    await page.getByRole("menuitemradio", { name: new RegExp(`^${target}`) }).focus();
    await page.keyboard.press("Enter");

    await page.locator(".pi-status-mode").filter({ hasText: target }).waitFor({ state: "visible", timeout: 1_000 });
    assert.equal(await page.evaluate(() => document.activeElement?.classList.contains("pi-status-mode") ?? false), true);
  });
});

void test("a tab switch closes the thinking popover instead of leaving it bound to the old tab", async () => {
  await withTaskpane({ width: 380, height: 720 }, async (page) => {
    await page.getByRole("button", { name: "New tab" }).click();
    await page.locator(".pi-session-tab").nth(1).waitFor({ state: "visible" });
    await page.locator(".pi-status-thinking").click();
    const popover = page.locator(".pi-status-popover--thinking");
    await popover.waitFor({ state: "visible" });
    // Keyboard switch: no outside click to close the popover as a side effect.
    await page.keyboard.press("Control+PageUp");
    await page.locator(".pi-session-tab").first().and(page.locator(".is-active")).waitFor({ state: "visible" });
    await popover.waitFor({ state: "detached", timeout: 2_000 });
  });
});

const bridgeRows = [
  { name: "tmux outage", tool: "python_transform_range", details: { kind: "tmux_bridge", ok: false, action: "capture_pane", error: "bridge unreachable", gateReason: "bridge_unreachable", skillHint: "tmux-bridge" }, title: /terminal access.*unavailable/i, enabled: true },
  { name: "tmux custom URL", tool: "python_transform_range", details: { kind: "tmux_bridge", ok: false, action: "list_sessions", bridgeUrl: "https://localhost:4441", error: "bridge unreachable", gateReason: "bridge_unreachable", skillHint: "tmux-bridge" }, title: /terminal access.*unavailable/i, enabled: true, probeUrl: "https://localhost:4441" },
  { name: "invalid tmux URL", tool: "python_transform_range", details: { kind: "tmux_bridge", ok: false, action: "list_sessions", error: "invalid bridge URL", gateReason: "invalid_bridge_url", skillHint: "tmux-bridge" }, title: /terminal access.*unavailable/i, enabled: false },
  { name: "missing Python runtime", tool: "python_transform_range", details: { kind: "python_bridge", ok: false, action: "run_python", error: "no_python_runtime", skillHint: "python-bridge" }, title: /python.*unavailable/i, enabled: true },
  { name: "LibreOffice outage", tool: "python_transform_range", details: { kind: "libreoffice_bridge", ok: false, action: "convert", bridgeUrl: "https://localhost:4450", error: "bridge unreachable", gateReason: "bridge_unreachable", skillHint: "python-bridge" }, title: /file conversion.*unavailable/i, enabled: true },
  { name: "ordinary Python error", tool: "python_transform_range", details: { kind: "python_bridge", ok: false, action: "run_python", error: "NameError: x is not defined" }, title: null, enabled: null },
  { name: "Python custom URL", tool: "python_transform_range", details: { kind: "python_bridge", ok: false, action: "run_python", bridgeUrl: "https://localhost:5540", error: "bridge unreachable", gateReason: "bridge_unreachable", skillHint: "python-bridge" }, title: /python bridge.*unavailable/i, enabled: true, probeUrl: "https://localhost:5540" },
] as const;

void test("the setup card renders the seven bridge-result states", async () => {
  await withTaskpane({ width: 420, height: 720 }, async (page) => {
    for (const row of bridgeRows) {
      await page.evaluate(async (details) => {
        document.getElementById("bridge-fixture")?.remove();
        delete document.body.dataset.probedBridgeUrl;
        const host = document.createElement("div");
        host.id = "bridge-fixture";
        document.body.appendChild(host);
        const [card, { decodeToolDetails }] = await Promise.all([
          import("/src/ui/bridge-setup-card.ts"),
          import("/src/tools/tool-details.ts"),
        ]);
        const decoded = decodeToolDetails(details);
        if (decoded && card.shouldShowBridgeSetupCard(decoded)) card.mountBridgeSetupCard(host, decoded);
      }, row.details);
      const setup = page.locator("#bridge-fixture .pi-bridge-setup");
      if (row.title === null) {
        assert.equal(await setup.count(), 0, row.name);
        continue;
      }
      await setup.waitFor({ state: "visible" });
      assert.match(await setup.locator(".pi-bridge-setup__title").innerText(), row.title, row.name);
      const button = setup.getByRole("button", { name: "Test connection" });
      assert.equal(await button.isEnabled(), row.enabled, row.name);
      if ("probeUrl" in row) {
        await page.evaluate(() => {
          Object.defineProperty(window, "fetch", { configurable: true, value: (input: RequestInfo | URL) => {
            document.body.dataset.probedBridgeUrl = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
            return Promise.resolve(new Response(JSON.stringify({ ok: true }), { status: 200, headers: { "content-type": "application/json" } }));
          } });
        });
        await button.click();
        await page.waitForFunction(() => document.body.dataset.probedBridgeUrl !== undefined);
        assert.match(await page.locator("body").getAttribute("data-probed-bridge-url") ?? "", new RegExp(row.probeUrl.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
      }
    }
  });
});
