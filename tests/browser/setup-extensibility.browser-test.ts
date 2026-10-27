import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { connectFixtureProvider, enterInput, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

void test("native search follows a selected turn model after a schema-stable model switch", async () => {
  const opened = await openTaskpane(env);
  const { page } = opened;
  try {
    let nativeRequest = "";
    await opened.context.route("**/api-proxy/openai/**", async (route) => {
      nativeRequest = route.request().postData() ?? "";
      await route.fulfill({ status: 400, contentType: "application/json", body: JSON.stringify({ error: { message: "Search fixture rejected", type: "invalid_request_error" } }) });
    });
    await connectFixtureProvider(page, "openai", "fixture-key");
    await page.locator(".pi-status-model").click();
    await page.getByPlaceholder("Search models…").fill("gpt-6-luna");
    await page.getByRole("option").filter({ hasText: "gpt-6-luna" }).click();
    await page.evaluate(async () => {
      const streamPath = "/node_modules/@earendil-works/pi-ai/dist/utils/event-stream.js";
      const fauxPath = "/node_modules/@earendil-works/pi-ai/dist/providers/faux.js";
      const sidebarPath = "/src/ui/pi-sidebar.ts";
      const { createAssistantMessageEventStream } = await import(/* @vite-ignore */ streamPath) as typeof import("@earendil-works/pi-ai/utils/event-stream");
      const { fauxAssistantMessage } = await import(/* @vite-ignore */ fauxPath) as typeof import("@earendil-works/pi-ai/providers/faux");
      const { PiSidebar } = await import(/* @vite-ignore */ sidebarPath) as typeof import("../../src/ui/pi-sidebar.ts");
      const sidebar = document.querySelector("pi-sidebar");
      if (!(sidebar instanceof PiSidebar) || !sidebar.agent) throw new Error("sidebar missing");
      let calls = 0;
      sidebar.agent.streamFunction = () => {
        const stream = createAssistantMessageEventStream();
        const result = ++calls === 1
          ? fauxAssistantMessage([{ type: "toolCall", id: "native-search", name: "web_search", arguments: { query: "official figures" } }], { stopReason: "toolUse" })
          : fauxAssistantMessage("Search fixture finished");
        stream.push({ type: "done", reason: result.stopReason === "toolUse" ? "toolUse" : "stop", message: result });
        return stream;
      };
    });
    await enterInput(page, "Search official figures");
    await page.getByRole("paragraph").filter({ hasText: "Search fixture finished" }).waitFor();
    assert.match(nativeRequest, /"model"\s*:\s*"gpt-6-luna"/);
    assert.match(nativeRequest, /"type"\s*:\s*"web_search"/);
  } finally { await opened.finish(); }
});

void test("no-provider send keeps the draft and opens labelled, no-setup-first provider choices", async () => {
  const opened = await openTaskpane(env);
  const { page } = opened;
  try {
    assert.equal(await page.locator(".pi-status-model__name").textContent(), "Connect a provider");
    await enterInput(page, "Keep this unsent question");
    const welcome = page.locator("#pi-welcome-login-overlay");
    await welcome.waitFor({ state: "visible" });
    assert.equal(await page.locator("pi-input textarea").inputValue(), "Keep this unsent question");
    const labels = await welcome.locator(".pi-login-label").allTextContents();
    assert.equal(labels[0], "OpenAI (API)");
    assert.ok(labels.indexOf("Google Gemini (API)") < labels.indexOf("Anthropic"));
    assert.match(await welcome.locator(".pi-login-row").first().innerText(), /Paste an API key: no setup/);
    assert.match(await welcome.locator(".pi-login-row").filter({ hasText: "Claude Pro/Max" }).innerText(), /Sign in: needs the helper app on this computer/);
  } finally { await opened.finish(); }
});

void test("composer drops keep renamed paths with the draft, remove chips, and send attachment intent", async () => {
  const opened = await openTaskpane(env);
  const { page } = opened;
  try {
    await connectFixtureProvider(page, "mistral", "fixture-key");
    await page.waitForFunction(() => !document.querySelector(".pi-status-model__name")?.textContent?.includes("Connect a provider"));
    for (const text of ["old", "new"]) {
      await page.locator("pi-input .pi-input-card").evaluate((node, text) => {
        const dataTransfer = new DataTransfer();
        dataTransfer.items.add(new File([text], "budget.csv", { type: "text/csv" }));
        node.dispatchEvent(new DragEvent("drop", { bubbles: true, dataTransfer }));
      }, text);
      await page.getByRole("button", { name: text === "old" ? "Remove attachment budget.csv" : "Remove attachment budget (2).csv", exact: true }).waitFor({ state: "visible" });
    }
    await page.getByRole("button", { name: "Remove attachment budget.csv", exact: true }).click();
    await page.locator("pi-input textarea").fill("Compare this with Sheet1");
    await page.waitForTimeout(650);
    await page.reload();
    await page.getByRole("button", { name: "Remove attachment budget (2).csv", exact: true }).waitFor({ state: "visible", timeout: 20_000 });
    assert.equal(await page.locator("pi-input textarea").inputValue(), "Compare this with Sheet1");
    assert.equal(await page.getByRole("button", { name: "Remove attachment budget.csv", exact: true }).count(), 0);
    await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar)) throw new Error("sidebar missing");
      const agent = sidebar.agent as { streamFunction: () => Promise<never>; state: { tools: Array<{ name: string }> } };
      assertNoSearch(agent.state.tools);
      agent.streamFunction = () => new Promise<never>(() => {});
      function assertNoSearch(tools: Array<{ name: string }>): void {
        if (tools.some((tool) => tool.name === "web_search")) throw new Error("unsupported/no-key search must not be offered");
      }
    });
    await page.locator("pi-input textarea").press("Enter");
    await page.locator("user-message").filter({ hasText: "Attached files in Files" }).waitFor({ state: "visible" });
    assert.match(await page.locator("user-message").last().innerText(), /budget \(2\).csv/);
    assert.equal(await page.locator(".pi-input-attachment").count(), 0);
  } finally { await opened.finish(); }
});

void test("Anthropic subscription warning is UI-only, once per session, and can be disabled", async () => {
  const opened = await openTaskpane(env);
  const { page } = opened;
  try {
    await connectFixtureProvider(page, "anthropic", "sk-ant-oat-fixture");
    const warning = page.locator(".pi-subscription-warning");
    await warning.waitFor({ state: "visible" });
    assert.equal(await warning.count(), 1);
    assert.match(await warning.innerText(), /billed per token, not your Claude plan limits/);
    assert.equal(await warning.locator("a").getAttribute("href"), "https://claude.ai/settings/usage");
    assert.equal(await page.evaluate(() => {
      const sidebar = document.querySelector("pi-sidebar");
      if (!sidebar || !("agent" in sidebar)) throw new Error("sidebar missing");
      return JSON.stringify((sidebar.agent as { state: { messages: object[] } }).state.messages).includes("Third-party harness");
    }), false);
    await page.reload();
    await warning.waitFor({ state: "visible", timeout: 20_000 });
    assert.equal(await warning.count(), 1);
    await enterInput(page, "/settings");
    const row = page.locator(".pi-set-row").filter({ hasText: "Anthropic subscription warning" });
    await row.locator(".pi-toggle").click();
    await page.locator("#pi-settings-overlay").dispatchEvent("click");
    await warning.waitFor({ state: "detached" });
  } finally { await opened.finish(); }
});
