/* oxlint-disable typescript/no-unsafe-assignment, typescript/no-unsafe-call, typescript/no-unsafe-member-access -- Playwright evaluates Vite-loaded application modules whose browser runtime types are unavailable to the Node test checker. */
import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import type { Page } from "playwright";

import { connectFixtureProvider, enterCommand, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, { viewport: { width: 420, height: 720 }, clientId: "ui-dialogs-browser" });
  try { await run(opened.page); } finally { await opened.finish(); }
}

async function importFile(page: Page, name: string, contents = "hello", type = "text/plain"): Promise<void> {
  await page.evaluate(async ({ name, contents, type }) => {
    const { getFilesWorkspace } = await import("/src/files/workspace.ts");
    await getFilesWorkspace().importFiles([new File([contents], name, { type, lastModified: 1 })]);
  }, { name, contents, type });
}

async function openFileDetail(page: Page, name: string): Promise<void> {
  await enterCommand(page, "/files");
  await page.getByText(name, { exact: true }).click();
  await page.locator(".pi-files-detail-actions").waitFor({ state: "visible" });
}

void test("export counts chat messages without discarding hidden debugging records", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(() => {
      const messages = document.querySelector("pi-sidebar").agent.state.messages;
      messages.splice(0, messages.length,
        { role: "autoContext", text: "Hidden workbook context", timestamp: 1 },
        { role: "user", content: "Visible request", timestamp: 2 },
        { role: "autoContext", text: "Updated workbook context", timestamp: 3 },
      );
    });
    const downloaded = page.waitForEvent("download");
    await enterCommand(page, "/export download");
    const stream = await (await downloaded).createReadStream();
    let json = "";
    for await (const chunk of stream) json += String(chunk);
    assert.partialDeepStrictEqual(JSON.parse(json), {
      messageCount: 1,
      raw: [
        { role: "autoContext", text: "Hidden workbook context" },
        { role: "user", content: "Visible request" },
        { role: "autoContext", text: "Updated workbook context" },
      ],
    });
  });
});

async function startMcpCall(page: Page, mode: "safe" | "yolo"): Promise<void> {
  await page.evaluate(async (mode) => {
    const { createMcpTool } = await import("/src/tools/mcp.ts") as typeof import("../../src/tools/mcp.ts");
    const { requestApproval } = await import("/src/ui/approval-dialog.ts") as typeof import("../../src/ui/approval-dialog.ts");
    document.body.dataset.mcpResult = "pending";
    const tool = createMcpTool({
      getExecutionMode: () => Promise.resolve(mode),
      requestToolApproval: (request) => requestApproval(request),
      getRuntimeConfig: () => Promise.resolve({ servers: [{ id: "consent", name: "Consent server", url: "https://localhost:4010/mcp", enabled: true }] }),
      callJsonRpc: ({ method }) => Promise.resolve({ proxied: false, result: method === "tools/list"
        ? { result: { tools: [{ name: "send_note" }] } }
        : { result: { content: [{ type: "text", text: "sent" }] } } }),
    });
    void tool.execute("call", { tool: "send_note", args: '{"note":"Consent test 17"}' }).then((result) => {
      document.body.dataset.mcpResult = result.details.ok ? "sent" : "declined";
    });
  }, mode);
}

void test("MCP Always allow persists, Confirm still asks, and Connections can revoke it", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const { saveMcpServers } = await import("/src/tools/mcp-config.ts") as typeof import("../../src/tools/mcp-config.ts");
      await saveMcpServers(getAppStorage().settings, [{ id: "consent", name: "Consent server", url: "https://localhost:4010/mcp", enabled: true }]);
    });
    await startMcpCall(page, "yolo");
    const dialog = page.locator("#pi-tool-approval-overlay");
    await dialog.getByText(/Consent test 17/).waitFor();
    await dialog.getByRole("button", { name: "Always allow this tool", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.mcpResult === "sent");
    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible" });
    const welcome = page.locator("#pi-welcome-login-overlay");
    await welcome.waitFor({ state: "visible" });
    await welcome.dispatchEvent("click");
    await welcome.waitFor({ state: "detached" });
    await startMcpCall(page, "yolo");
    await page.waitForFunction(() => document.body.dataset.mcpResult === "sent");
    assert.equal(await dialog.count(), 0);
    await startMcpCall(page, "safe");
    await dialog.waitFor({ state: "visible" });
    assert.equal(await dialog.getByRole("button", { name: "Always allow this tool", exact: true }).count(), 0);
    assert.equal(await dialog.getByRole("button", { name: "Allow for the rest of this turn", exact: true }).count(), 0);
    await dialog.getByRole("button", { name: "Cancel", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.mcpResult === "declined");
    await enterCommand(page, "/tools");
    await page.getByText("Consent server", { exact: true }).click();
    await page.getByRole("button", { name: "Revoke “send_note”", exact: true }).click();
    await page.getByText("Revoked approval for “send_note”", { exact: true }).waitFor();
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await startMcpCall(page, "yolo");
    await dialog.waitFor({ state: "visible" });
    await dialog.getByRole("button", { name: "Cancel", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.mcpResult === "declined");
  });
});

void test("guidance tools preview Confirm changes and Undo restores rules, conventions and installed skills", async () => {
  await withTaskpane(async (page) => {
    const welcome = page.locator("#pi-welcome-login-overlay");
    if (await welcome.count()) await welcome.dispatchEvent("click");
    for (const kind of ["rules", "conventions", "skills"] as const) {
      await page.evaluate(async (kind) => {
        const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
        const { requestApproval } = await import("/src/ui/approval-dialog.ts") as typeof import("../../src/ui/approval-dialog.ts");
        const { createInstructionsTool } = await import("/src/tools/instructions.ts");
        const { createConventionsTool } = await import("/src/tools/conventions.ts");
        const { createSkillsTool } = await import("/src/tools/skills.ts");
        const { setUserRules } = await import("/src/rules/store.ts");
        const { setStoredConventions } = await import("/src/conventions/store.ts");
        const settings = getAppStorage().settings;
        const dependencies: import("../../src/tools/guidance-change.ts").GuidanceChangeDependencies = {
          getExecutionMode: () => Promise.resolve("safe" as const),
          requestLocalChangeApproval: (request) => requestApproval(request),
        };
        document.body.dataset.guidanceResult = "pending";
        if (kind === "rules") {
          await setUserRules(settings, "Keep the original rule");
          void createInstructionsTool(dependencies).execute("rule", { action: "append", level: "user", content: "Use 17 decimal places" }).then(() => { document.body.dataset.guidanceResult = "saved"; });
        } else if (kind === "conventions") {
          await setStoredConventions(settings, { visualDefaults: { fontName: "Arial", fontSize: 13 } });
          void createConventionsTool(dependencies).execute("convention", { action: "reset" }).then(() => { document.body.dataset.guidanceResult = "saved"; });
        } else {
          void createSkillsTool(dependencies).execute("skill", { action: "install", name: "consent-test", markdown: "---\nname: consent-test\ndescription: Consent test\n---\nUse 17 decimal places" }).then(() => { document.body.dataset.guidanceResult = "saved"; });
        }
      }, kind);
      const dialog = page.locator("#pi-tool-approval-overlay");
      await dialog.waitFor({ state: "visible" });
      assert.match(await dialog.innerText(), kind === "conventions" ? /New formatting conventions/ : /17 decimal places/);
      assert.equal(await dialog.getByRole("button", { name: "Allow for the rest of this turn", exact: true }).count(), 0);
      await dialog.getByRole("button", { name: "Save change", exact: true }).click();
      await page.waitForFunction(() => document.body.dataset.guidanceResult === "saved");
      await page.getByRole("button", { name: "Undo", exact: true }).click();
      await page.getByText("Undid Pi’s guidance change", { exact: true }).waitFor();
      const state = await page.evaluate(async (kind) => {
        const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
        const settings = getAppStorage().settings;
        if (kind === "rules") return settings.get("user.instructions");
        if (kind === "conventions") {
          const { getStoredConventions } = await import("/src/conventions/store.ts") as typeof import("../../src/conventions/store.ts");
          return getStoredConventions(settings);
        }
        const { getFilesWorkspace } = await import("/src/files/workspace.ts") as typeof import("../../src/files/workspace.ts");
        return (await getFilesWorkspace().listFiles()).some((file) => file.path === "skills/external/consent-test/SKILL.md");
      }, kind);
      assert.deepEqual(state, kind === "rules" ? "Keep the original rule" : kind === "conventions" ? { visualDefaults: { fontName: "Arial", fontSize: 13 } } : false);
      await page.mouse.move(5, 5);
      await page.getByText("Undid Pi’s guidance change", { exact: true }).waitFor({ state: "detached" });
    }
  });
});

void test("workbook rules ask even in Auto, and declining leaves the standing rule unchanged", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const { createInstructionsTool } = await import("/src/tools/instructions.ts") as typeof import("../../src/tools/instructions.ts");
      const { requestApproval } = await import("/src/ui/approval-dialog.ts") as typeof import("../../src/ui/approval-dialog.ts");
      const { BrowserHost } = await import("/src/host/browser-host.ts") as typeof import("../../src/host/browser-host.ts");
      const { getCurrentSpreadsheetHost, setCurrentSpreadsheetHost } = await import("/src/host/current.ts") as typeof import("../../src/host/current.ts");
      const { setWorkbookRules } = await import("/src/rules/store.ts") as typeof import("../../src/rules/store.ts");
      const previousHost = getCurrentSpreadsheetHost();
      class IdentifiedHost extends BrowserHost {
        override getWorkbookContext(): Promise<import("../../src/workbook/context.ts").WorkbookContext> {
          return Promise.resolve({ workbookId: "consent-workbook", workbookName: "Consent.xlsx", source: "document.url" });
        }
      }
      setCurrentSpreadsheetHost(new IdentifiedHost());
      await setWorkbookRules(getAppStorage().settings, "consent-workbook", "Original workbook rule");
      document.body.dataset.workbookRuleResult = "pending";
      void createInstructionsTool({
        getExecutionMode: () => Promise.resolve("yolo"),
        requestLocalChangeApproval: (request) => requestApproval(request),
      }).execute("rule", { action: "replace", level: "workbook", content: "Replace all standing workbook guidance" }).then(() => {
        document.body.dataset.workbookRuleResult = "settled";
      }).finally(() => setCurrentSpreadsheetHost(previousHost));
    });
    const dialog = page.locator("#pi-tool-approval-overlay");
    await dialog.getByText(/New rules for this workbook/).waitFor();
    await dialog.getByRole("button", { name: "Cancel", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.workbookRuleResult === "settled");
    assert.equal(await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      return getAppStorage().settings.get("workbook.instructions.v1.consent-workbook");
    }), "Original workbook rule");
  });
});

void test("Stop Pi in a local-file approval aborts the turn without replacing the user's file", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts");
      await getAppStorage().settings.set("execution.mode.v1", "safe");
      const { getFilesWorkspace } = await import("/src/files/workspace.ts");
      await getFilesWorkspace().writeTextFile("stop-approval.txt", "keep this", "text/plain");
    });
    await page.reload();
    await page.locator("pi-input textarea").waitFor({ state: "visible", timeout: 20_000 });
    await connectFixtureProvider(page);
    await page.evaluate(async () => {
      const { createAssistantMessageEventStream } = await import("/node_modules/@earendil-works/pi-ai/dist/utils/event-stream.js") as typeof import("@earendil-works/pi-ai/utils/event-stream");
      const { fauxAssistantMessage } = await import("/node_modules/@earendil-works/pi-ai/dist/providers/faux.js") as typeof import("@earendil-works/pi-ai/providers/faux");
      const { PiSidebar } = await import("/src/ui/pi-sidebar.ts") as typeof import("../../src/ui/pi-sidebar.ts");
      const sidebar = document.querySelector("pi-sidebar");
      if (!(sidebar instanceof PiSidebar) || !sidebar.agent) throw new Error("sidebar missing");
      sidebar.agent.getApiKey = () => Promise.resolve("synthetic-key");
      sidebar.agent.streamFunction = (_model, _context, options) => {
        const stream = createAssistantMessageEventStream();
        if (options?.signal?.aborted) {
          stream.push({ type: "error", reason: "aborted", error: fauxAssistantMessage("", { stopReason: "aborted" }) });
        } else {
          stream.push({ type: "done", reason: "toolUse", message: fauxAssistantMessage([
            { type: "toolCall", id: "create-file", name: "files", arguments: { action: "write", path: "stop-approval.txt", content: "must not be written" } },
          ], { stopReason: "toolUse" }) });
        }
        return stream;
      };
    });
    const welcome = page.locator("#pi-welcome-login-overlay");
    if (await welcome.count()) await welcome.dispatchEvent("click");
    await enterCommand(page, "Create the file");
    const dialog = page.locator("#pi-tool-approval-overlay");
    await dialog.waitFor({ state: "visible" });
    await page.waitForFunction(() => document.querySelector(".pi-announcer")?.textContent?.includes("approval"));
    await dialog.getByRole("button", { name: "Stop Pi", exact: true }).click({ timeout: 3_000 });
    await dialog.waitFor({ state: "detached" });
    await page.waitForFunction(() => !document.querySelector("pi-sidebar").agent.state.isStreaming);
    await page.waitForTimeout(100);
    assert.equal(await page.locator(".pi-announcer").innerText(), "");
    const contents = await page.evaluate(async () => {
      const { getFilesWorkspace } = await import("/src/files/workspace.ts") as typeof import("../../src/files/workspace.ts");
      return (await getFilesWorkspace().readFile("stop-approval.txt", { mode: "text" })).text;
    });
    assert.equal(contents, "keep this");
    assert.equal(await page.locator("[inert]").count(), 0);
  });
});

void test("withdrawal cancels delayed approval text and preserves a newer announcement", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { requestApproval } = await import("/src/ui/approval-dialog.ts");
      const stopped = new AbortController();
      const answer = requestApproval({ title: "Withdraw immediately", confirmLabel: "Allow", signal: stopped.signal });
      stopped.abort();
      await answer;
    });
    await page.waitForTimeout(100);
    assert.equal(await page.locator(".pi-announcer").innerText(), "");
    await page.evaluate(async () => {
      const { requestApproval } = await import("/src/ui/approval-dialog.ts");
      const { announce } = await import("/src/ui/announcer.ts");
      void requestApproval({ title: "Old approval", confirmLabel: "Allow" });
      announce("Another tab finished");
    });
    await page.waitForFunction(() => document.querySelector(".pi-announcer")?.textContent === "Another tab finished");
    await page.locator("#pi-tool-approval-overlay").getByRole("button", { name: "Decline", exact: true }).click();
    assert.equal(await page.locator(".pi-announcer").innerText(), "Another tab finished");
  });
});

void test("approvals ignore a stray Enter, and Escape declines only the shown request", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { requestApproval } = await import("/src/ui/approval-dialog.ts");
      document.body.dataset.approvalOutcomes = "";
      for (const title of ["Write D1", "Write D2"]) {
        void requestApproval({ title, confirmLabel: "Allow" }).then((approved) => {
          document.body.dataset.approvalOutcomes += `${title}:${approved};`;
        });
      }
    });
    const dialog = page.locator("#pi-tool-approval-overlay");
    await dialog.getByRole("heading", { name: "Write D1" }).waitFor();
    await page.waitForFunction(() => document.querySelector(".pi-announcer")?.textContent?.includes("Write D1"));
    await page.waitForFunction(() => Boolean(document.activeElement?.closest("#pi-tool-approval-overlay")));
    assert.equal(await dialog.getByRole("button", { name: "Allow", exact: true }).evaluate((button) => button === document.activeElement), false);
    await page.keyboard.press("Enter");
    assert.equal(await page.locator("body").getAttribute("data-approval-outcomes"), "");
    await page.keyboard.press("Escape");
    await dialog.getByRole("heading", { name: "Write D2" }).waitFor();
    await page.waitForFunction(() => document.querySelector(".pi-announcer")?.textContent?.includes("Write D2"));
    assert.equal(await page.locator("body").getAttribute("data-approval-outcomes"), "Write D1:false;");
    await dialog.getByRole("button", { name: "Allow", exact: true }).click();
    await dialog.waitFor({ state: "detached" });
    await page.waitForTimeout(100);
    assert.equal(await page.locator(".pi-announcer").innerText(), "");
    assert.equal(await page.locator("body").getAttribute("data-approval-outcomes"), "Write D1:false;Write D2:true;");
  });
});

void test("the Files rename dialog confirms a new path", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "plan.md");
    await openFileDetail(page, "plan.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    await dialog.getByRole("textbox").fill("archive/final-plan");
    await dialog.getByRole("button", { name: "Rename" }).click();
    await page.getByRole("heading", { name: "final-plan.md" }).waitFor({ state: "visible" });
    assert.equal(await page.getByRole("heading", { name: "final-plan.md" }).innerText(), "final-plan.md");
    await page.waitForTimeout(200);
    const notices = await page.locator("#pi-toast .pi-toast__message").allTextContents();
    assert.deepEqual(notices.filter((text) => text.startsWith("Renamed to") || text === "That file is no longer available."),
      ["Renamed to archive/final-plan.md."]);
  });
});

void test("a Files preview delete reports success, while an external delete still warns", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "delete-me.md");
    await openFileDetail(page, "delete-me.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Delete", exact: true }).click();
    await page.getByRole("dialog").last().getByRole("button", { name: "Delete", exact: true }).click();
    await page.getByRole("heading", { name: "Files", exact: true }).waitFor();
    await page.waitForTimeout(200);
    const notices = await page.locator("#pi-toast .pi-toast__message").allTextContents();
    assert.deepEqual(notices.filter((text) => text.startsWith("Deleted") || text === "That file is no longer available."),
      ["Deleted delete-me.md."]);

    await importFile(page, "external-delete.md");
    await page.getByText("external-delete.md", { exact: true }).click();
    await page.getByRole("heading", { name: "external-delete.md", exact: true }).waitFor();
    await page.evaluate(async () => {
      const { getFilesWorkspace } = await import("/src/files/workspace.ts");
      await getFilesWorkspace().deleteFile("external-delete.md", { audit: { actor: "assistant", source: "tool:files" } });
    });
    await page.getByText("That file is no longer available.", { exact: true }).waitFor();
    await page.getByRole("heading", { name: "Files", exact: true }).waitFor();
  });
});

void test("the Files rename dialog cancels without changing the file", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "cancel-me.md");
    await openFileDetail(page, "cancel-me.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    await dialog.getByRole("textbox").fill("changed.md");
    await dialog.getByRole("button", { name: "Cancel", exact: true }).last().click();
    await dialog.waitFor({ state: "detached" });
    assert.equal(await page.getByRole("heading", { name: "cancel-me.md" }).innerText(), "cancel-me.md");
  });
});

void test("Escape leaves text entry then dismisses the dialog and permits a fresh mount", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "escape.md");
    await openFileDetail(page, "escape.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    await page.waitForFunction(() => document.activeElement?.matches("#pi-text-input-dialog-overlay input") === true);
    await page.keyboard.press("Escape");
    await page.keyboard.press("Escape");
    await dialog.waitFor({ state: "detached" });
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    assert.equal(await page.locator("#pi-text-input-dialog-overlay").count(), 1);
  });
});

void test("an overlay exposes modal dialog semantics and focuses its text input", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "focus.md");
    await openFileDetail(page, "focus.md");
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
    const dialog = page.locator("#pi-text-input-dialog-overlay");
    // The dialog moves focus on the next animation frame after opening.
    await page.waitForFunction(
      () => document.activeElement?.closest("#pi-text-input-dialog-overlay") !== null,
      undefined,
      { timeout: 5_000 },
    );
    assert.deepEqual({
      role: await dialog.getAttribute("role"),
      modal: await dialog.getAttribute("aria-modal"),
      focused: await dialog.getByRole("textbox").evaluate((input) => input === document.activeElement),
    }, { role: "dialog", modal: "true", focused: true });
  });
});

void test("read-only built-in files expose only copy and download actions", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/files");
    const builtIn = page.locator(".pi-files-item--muted").first();
    await builtIn.click();
    const actions = page.locator(".pi-files-detail-actions button");
    await actions.last().waitFor({ state: "visible" });
    assert.deepEqual(await actions.allTextContents(), ["Copy content", "Download"]);
  });
});

for (const row of [
  { name: "unsafe.html", type: "text/html", expected: "application/octet-stream" },
  { name: "safe.txt", type: "text/plain", expected: "text/plain" },
] as const) {
  void test(`Files Open uses a safe blob type for ${row.type}`, async () => {
    await withTaskpane(async (page) => {
      await importFile(page, row.name, "sample", row.type);
      await page.evaluate(() => {
        Object.defineProperty(window, "open", { configurable: true, value: () => null });
        const original = URL.createObjectURL.bind(URL);
        Object.defineProperty(URL, "createObjectURL", { configurable: true, value: (blob: Blob) => {
          document.body.dataset.openedBlobType = blob.type;
          return original(blob);
        } });
      });
      await openFileDetail(page, row.name);
      await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Open ↗" }).evaluate((button) => button.dispatchEvent(new MouseEvent("click", { bubbles: true })));
      await page.waitForFunction(() => document.body.dataset.openedBlobType !== undefined);
      assert.equal(await page.locator("body").getAttribute("data-opened-blob-type"), row.expected);
    });
  });
}

void test("the Files footer reports count, total size, and active backend", async () => {
  await withTaskpane(async (page) => {
    await importFile(page, "sized.txt", "x".repeat(1024));
    await enterCommand(page, "/files");
    const footer = page.locator(".pi-files-footer");
    await page.waitForFunction(() => (document.querySelector(".pi-files-footer")?.textContent?.length ?? 0) > 0);
    assert.match(await footer.innerText(), /\d+ files? · [\d.]+ (?:KB|B) · Sandboxed workspace/);
  });
});
