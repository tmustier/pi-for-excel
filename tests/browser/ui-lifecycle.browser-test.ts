import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import type { Page } from "playwright";
import { enterCommand, openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

async function withTaskpane(run: (page: Page) => Promise<void>): Promise<void> {
  const opened = await openTaskpane(env, { clientId: "ui-lifecycle" });
  try { await run(opened.page); } finally { await opened.finish(); }
}

async function release(page: Page, event = "fixture-release"): Promise<void> {
  await page.evaluate((name) => { document.dispatchEvent(new Event(name)); }, event);
}

void test("inline approval migration keeps a single preparation owner", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const api = await import("/src/ui/approval-dialog.ts") as typeof import("../../src/ui/approval-dialog.ts");
      const prepare = document.createElement("button");
      prepare.textContent = "Prepare fixture";
      document.body.append(prepare);
      const host = api.registerInlineApprovalHost({ claims: () => true, onChange: () => {} });
      prepare.onclick = () => { void api.getInlineApprovals()[0]?.prepareAndApprove(); host.unregister(); };
      void api.requestApproval({ title: "Migrating approval", toolCallId: "fixture", confirmLabel: "Allow", secondaryApproval: {
        label: "Prepare", onApprove: () => new Promise<void>((resolve) => { document.addEventListener("fixture-release", () => resolve(), { once: true }); }),
      } }).then((answer) => { document.body.dataset.approval = String(answer); });
    });
    await page.getByRole("button", { name: "Prepare fixture", exact: true }).click();
    assert.equal(await page.locator("#pi-tool-approval-overlay").count(), 0, "migration must not offer a second preparation");
    await release(page);
    await page.waitForFunction(() => document.body.dataset.approval === "true");
  });
});

void test("an abort that migrates an inline approval is settled rather than queued", async () => {
  await withTaskpane(async (page) => {
    const result = await page.evaluate(async () => {
      const api = await import("/src/ui/approval-dialog.ts") as typeof import("../../src/ui/approval-dialog.ts");
      const controller = new AbortController();
      const host = api.registerInlineApprovalHost({ claims: () => true, onChange: () => {} });
      controller.signal.addEventListener("abort", () => host.unregister(), { once: true });
      let outcome: boolean | null = null;
      void api.requestApproval({ title: "Aborted migration", toolCallId: "fixture", confirmLabel: "Allow", signal: controller.signal }).then((answer) => { outcome = answer; });
      controller.abort();
      await Promise.resolve();
      return { outcome, dialogs: document.querySelectorAll("#pi-tool-approval-overlay").length };
    });
    assert.deepEqual(result, { outcome: false, dialogs: 0 });
  });
});

void test("Rules saves are exclusive and preserve edits entered during persistence", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/rules");
    const rules = page.getByPlaceholder(/Your preferences and habits/);
    await rules.fill("First draft");
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const settings = getAppStorage().settings;
      const set = settings.set.bind(settings);
      let writes = 0;
      settings.set = async (...args) => {
        if (args[0] === "user.instructions") {
          document.body.dataset.ruleWrites = String(++writes);
          if (writes === 1) await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
        }
        await set(...args);
      };
      const save = document.querySelector<HTMLButtonElement>(".pi-rules-page__footer .pi-overlay-btn--primary");
      save?.dispatchEvent(new MouseEvent("click"));
      save?.dispatchEvent(new MouseEvent("click"));
    });
    await page.waitForFunction(() => document.body.dataset.ruleWrites !== undefined);
    assert.equal(await page.locator("body").getAttribute("data-rule-writes"), "1", "only one Rules save may write at a time");
    await rules.fill("Newer draft");
    await release(page);
    await page.locator(".pi-rules-page__footer .pi-overlay-btn--primary:not([disabled])").waitFor();
    assert.equal(await rules.inputValue(), "Newer draft");
    await page.locator(".pi-rules-page__footer .pi-overlay-btn--primary").click();
    await page.locator("#pi-settings-overlay").waitFor({ state: "detached" });
    assert.equal(await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      return getAppStorage().settings.get("user.instructions");
    }), "Newer draft");
  });
});

void test("Rules save failures are visible and allow retry", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/rules");
    await page.getByPlaceholder(/Your preferences and habits/).fill("Retry draft");
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const settings = getAppStorage().settings;
      const set = settings.set.bind(settings);
      settings.set = async (...args) => {
        if (args[0] === "user.instructions") { settings.set = set; throw new Error("fixture unavailable"); }
        await set(...args);
      };
    });
    await page.locator(".pi-rules-page__footer .pi-overlay-btn--primary").click();
    await page.getByText("Save failed: fixture unavailable", { exact: true }).waitFor();
    await page.locator(".pi-rules-page__footer .pi-overlay-btn--primary").click();
    await page.locator("#pi-settings-overlay").waitFor({ state: "detached" });
  });
});

void test("closing Files while an image is loading revokes the late preview URL", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getFilesWorkspace } = await import("/src/files/workspace.ts") as typeof import("../../src/files/workspace.ts");
      const workspace = getFilesWorkspace();
      await workspace.importFiles([new File([new Uint8Array([1, 2, 3])], "late.png", { type: "image/png" })]);
      const read = workspace.readFile.bind(workspace);
      workspace.readFile = async (...args) => {
        const result = await read(...args);
        if (args[0] === "late.png") {
          document.body.dataset.previewLoading = "true";
          await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
        }
        return result;
      };
      const create = URL.createObjectURL.bind(URL);
      const revoke = URL.revokeObjectURL.bind(URL);
      URL.createObjectURL = (blob) => { const url = create(blob); document.body.dataset.createdUrl = url; return url; };
      URL.revokeObjectURL = (url) => { document.body.dataset.revokedUrl = url; revoke(url); };
    });
    await enterCommand(page, "/files");
    await page.getByText("late.png", { exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.previewLoading === "true");
    await page.locator("#pi-files-workspace-overlay .pi-overlay-header:not([hidden]) .pi-overlay-close").click();
    await release(page);
    await page.waitForFunction(() => document.body.dataset.createdUrl !== undefined);
    assert.equal(await page.locator("body").getAttribute("data-revoked-url"), await page.locator("body").getAttribute("data-created-url"));
  });
});

void test("a completed rename does not reopen Files detail after Back", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getFilesWorkspace } = await import("/src/files/workspace.ts") as typeof import("../../src/files/workspace.ts");
      const workspace = getFilesWorkspace();
      await workspace.writeTextFile("old.md", "fixture", "text/markdown");
      const rename = workspace.renameFile.bind(workspace);
      workspace.renameFile = async (...args) => {
        document.body.dataset.renaming = "true";
        await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
        await rename(...args);
      };
    });
    await enterCommand(page, "/files");
    await page.getByText("old.md", { exact: true }).click();
    await page.locator(".pi-files-detail-actions").getByRole("button", { name: "Rename", exact: true }).click();
    const prompt = page.locator("#pi-text-input-dialog-overlay");
    await prompt.getByRole("textbox").fill("new.md");
    await prompt.getByRole("button", { name: "Rename", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.renaming === "true");
    await page.locator(".pi-files-detail__back").click();
    await release(page);
    await page.locator(".pi-files-item__name").filter({ hasText: "new.md" }).waitFor({ state: "attached" });
    assert.equal(await page.getByRole("heading", { name: "Files", exact: true }).isVisible(), true);
  });
});

void test("Settings discards stale concurrent navigation requests", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { createSettingsShell } = await import("/src/ui/settings-shell.ts") as typeof import("../../src/ui/settings-shell.ts");
      const shell = createSettingsShell({ overlayId: "fixture-settings", rootId: "root", backLabel: () => "Back fixture", closeLabel: () => "Close fixture", getPage: (id) => ({
        id, title: () => id, render: (ctx) => {
          if (id !== "root") return;
          const button = document.createElement("button");
          button.textContent = "Navigate twice";
          button.onclick = () => { ctx.navigate("first"); ctx.navigate("second"); };
          ctx.body.append(button);
        },
      }) });
      await shell.open();
    });
    await page.getByRole("button", { name: "Navigate twice" }).click();
    assert.equal(await page.locator("#fixture-settings .pi-overlay-title").innerText(), "first");
    await page.getByRole("button", { name: "Back fixture" }).click();
    await page.getByRole("button", { name: "Navigate twice" }).waitFor();
  });
});

void test("proxy toggle persistence is serialized and an old failure cannot roll back a newer toggle", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/settings");
    await page.getByRole("button", { name: /^Proxy/ }).click();
    await page.locator(".pi-settings-proxy-url").waitFor();
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const settings = getAppStorage().settings;
      const set = settings.set.bind(settings);
      let writes = 0;
      settings.set = async (...args) => {
        if (args[0] === "proxy.enabled") {
          document.body.dataset.proxyWrites = String(++writes);
          if (writes === 1) {
            await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
            throw new Error("old failure");
          }
        }
        await set(...args);
      };
      const toggle = document.querySelector<HTMLInputElement>(".pi-settings-proxy-toggle input");
      if (!toggle) throw new Error("toggle missing");
      for (const checked of [true, false]) { toggle.checked = checked; toggle.dispatchEvent(new Event("change")); }
    });
    await page.waitForFunction(() => document.body.dataset.proxyWrites !== undefined);
    assert.equal(await page.locator("body").getAttribute("data-proxy-writes"), "1");
    await release(page);
    await page.waitForFunction(() => document.body.dataset.proxyWrites === "2");
    assert.equal(await page.locator(".pi-settings-proxy-toggle input").isChecked(), false);
  });
});

void test("gateway persistence locks Edit, Delete and form inputs until the original save completes", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const { saveOpenAiGatewayConfig } = await import("/src/auth/custom-gateways.ts") as typeof import("../../src/auth/custom-gateways.ts");
      await saveOpenAiGatewayConfig(getAppStorage().customProviders, { displayName: "Existing fixture", endpointUrl: "https://gateway.example.com/v1", modelId: "fixture" });
    });
    await enterCommand(page, "/settings");
    await page.getByRole("button", { name: /^Custom gateway/ }).click();
    await page.getByPlaceholder("Gateway name (optional)").fill("New fixture");
    await page.getByPlaceholder("https://your-gateway.example.com/v1").fill("https://new.example.com/v1");
    await page.getByPlaceholder("model-id").fill("new-model");
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const providers = getAppStorage().customProviders;
      const set = providers.set.bind(providers);
      providers.set = async (...args) => {
        document.body.dataset.gatewaySaving = "true";
        await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
        await set(...args);
      };
    });
    await page.getByRole("button", { name: "Save gateway", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.gatewaySaving === "true");
    assert.equal(await page.getByRole("button", { name: "Edit", exact: true }).isDisabled(), true);
    assert.equal(await page.getByRole("button", { name: "Delete", exact: true }).isDisabled(), true);
    assert.equal(await page.getByPlaceholder("Gateway name (optional)").isDisabled(), true);
    await release(page);
    await page.getByText("New fixture", { exact: true }).waitFor();
    await page.getByRole("button", { name: "Save gateway", exact: true }).waitFor({ state: "visible" });
  });
});

void test("Resume dismissal during a session read never opens that session", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const { PiSidebar } = await import("/src/ui/pi-sidebar.ts") as typeof import("../../src/ui/pi-sidebar.ts");
      const { showResumeDialog } = await import("/src/commands/builtins/resume-overlay.ts") as typeof import("../../src/commands/builtins/resume-overlay.ts");
      const sidebar = document.querySelector("pi-sidebar");
      if (!(sidebar instanceof PiSidebar) || !sidebar.agent) throw new Error("agent missing");
      const storage = getAppStorage();
      const now = new Date().toISOString();
      const data: import("../../src/storage/local/types.ts").SessionData = { id: "resume-fixture", title: "Saved fixture", model: { ...sidebar.agent.state.model, api: "openai-responses", provider: "openai" }, thinkingLevel: "off", messages: [{ role: "user", content: "fixture", timestamp: 1 }], createdAt: now, lastModified: now };
      await storage.sessions.save(data, { id: data.id, title: data.title, createdAt: now, lastModified: now, messageCount: 1, thinkingLevel: "off", preview: "fixture", usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0, cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } } });
      const load = storage.sessions.loadSession.bind(storage.sessions);
      storage.sessions.loadSession = async (id) => {
        document.body.dataset.sessionLoading = "true";
        await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
        const result = await load(id);
        document.body.dataset.sessionRead = "true";
        return result;
      };
      await showResumeDialog({ getRecentlyClosedItems: () => [], onReopenRecentlyClosed: () => Promise.resolve(false), onOpenSession: () => { document.body.dataset.openedSession = "true"; return Promise.resolve({ outcome: "opened", title: "Saved fixture" }); } });
    });
    await page.getByText("Saved fixture", { exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.sessionLoading === "true");
    await page.keyboard.press("Escape");
    await release(page);
    await page.waitForFunction(() => document.body.dataset.sessionRead === "true");
    assert.equal(await page.locator("body").getAttribute("data-opened-session"), null);
  });
});

void test("sidebar reconnects to the same agent after a disconnect", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { PiSidebar } = await import("/src/ui/pi-sidebar.ts") as typeof import("../../src/ui/pi-sidebar.ts");
      const sidebar = document.querySelector("pi-sidebar");
      if (!(sidebar instanceof PiSidebar) || !sidebar.agent) throw new Error("sidebar missing");
      const agent = sidebar.agent;
      let emit: (event: import("@earendil-works/pi-agent-core").AgentEvent) => void = () => {};
      agent.subscribe = (listener) => { emit = (event) => { void listener(event); }; return () => { emit = () => {}; }; };
      sidebar.remove();
      document.body.append(sidebar);
      await sidebar.updateComplete;
      agent.state.isStreaming = true;
      emit({ type: "agent_start" });
      await sidebar.updateComplete;
    });
    assert.equal(await page.getByRole("button", { name: "Stop", exact: true }).isVisible(), true);
  });
});

void test("toast focus and hover pause independently and duplicate refresh preserves the pause", async () => {
  await withTaskpane(async (page) => {
    await page.clock.install();
    await page.evaluate(async () => {
      const { showToast } = await import("/src/ui/toast.ts") as typeof import("../../src/ui/toast.ts");
      showToast("Paused error", { variant: "error", duration: 1000 });
    });
    const toast = page.locator(".pi-toast").filter({ hasText: "Paused error" });
    await toast.hover();
    await toast.getByRole("button").focus();
    await page.mouse.move(1, 1);
    await page.clock.runFor(1100);
    assert.equal(await toast.count(), 1, "pointerleave must not resume a focused toast");
    await page.evaluate(async () => {
      const { showToast } = await import("/src/ui/toast.ts") as typeof import("../../src/ui/toast.ts");
      showToast("Paused error", { variant: "error", duration: 1000 });
    });
    await page.clock.runFor(1100);
    assert.equal(await toast.count(), 1, "a duplicate refresh must remain paused");
    await page.locator("pi-input textarea").focus();
    await page.clock.runFor(1100);
    assert.equal(await toast.count(), 0);
  });
});

void test("toast duplicate matching keeps distinct supplied ids dismissible", async () => {
  await withTaskpane(async (page) => {
    const messages = await page.evaluate(async () => {
      const { showToast, dismissToast } = await import("/src/ui/toast.ts") as typeof import("../../src/ui/toast.ts");
      showToast("Same result", { id: "first", duration: 60_000 });
      showToast("Same result", { id: "second", duration: 60_000 });
      dismissToast("first");
      return [...document.querySelectorAll(".pi-toast__message")].map((node) => node.textContent);
    });
    assert.deepEqual(messages, ["Same result"]);
  });
});

void test("proxy banner ignores an older unavailable result after detection", async () => {
  await withTaskpane(async (page) => {
    const count = await page.evaluate(async () => {
      const { createProxyBannerController } = await import("/src/ui/proxy-banner.ts") as typeof import("../../src/ui/proxy-banner.ts");
      let resolveEnabled: (enabled: boolean) => void = () => {};
      const controller = createProxyBannerController({ isProxyEnabled: () => new Promise<boolean>((resolve) => { resolveEnabled = resolve; }) });
      const older = controller.update("not-detected");
      await controller.update("detected");
      resolveEnabled(true);
      await older;
      return document.querySelectorAll('[data-banner-id="proxy"]').length;
    });
    assert.equal(count, 0);
  });
});

void test("closing a covered middle overlay keeps the oldest inert until the top closes", async () => {
  await withTaskpane(async (page) => {
    const states = await page.evaluate(async () => {
      const { createOverlayDialog } = await import("/src/ui/overlay-dialog.ts") as typeof import("../../src/ui/overlay-dialog.ts");
      const dialogs = ["oldest", "middle", "top"].map((overlayId) => createOverlayDialog({ overlayId, cardClassName: "pi-overlay-card", restoreFocusOnClose: false }));
      for (const dialog of dialogs) dialog.mount();
      const [oldest, middle, top] = dialogs;
      if (!oldest || !middle || !top) throw new Error("missing fixture");
      middle.close();
      const covered = oldest.overlay.inert;
      top.close();
      const uncovered = oldest.overlay.inert;
      oldest.close();
      return { covered, uncovered };
    });
    assert.deepEqual(states, { covered: true, uncovered: false });
  });
});

void test("overlay manager ensure before mount shares one controller and dismisses it", async () => {
  await withTaskpane(async (page) => {
    const state = await page.evaluate(async () => {
      const { createOverlayDialogManager } = await import("/src/ui/overlay-dialog.ts") as typeof import("../../src/ui/overlay-dialog.ts");
      const manager = createOverlayDialogManager({ overlayId: "managed-fixture", cardClassName: "pi-overlay-card" });
      const first = manager.ensure();
      const second = manager.ensure();
      first.mount();
      second.mount();
      const count = document.querySelectorAll("#managed-fixture").length;
      manager.dismiss();
      return { count, remaining: document.querySelectorAll("#managed-fixture").length };
    });
    assert.deepEqual(state, { count: 1, remaining: 0 });
  });
});

void test("a failed search setup detection can be mounted again", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const { setAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const { mountSearchSetupCard } = await import("/src/ui/web-search-setup-card.ts") as typeof import("../../src/ui/web-search-setup-card.ts");
      const storage = getAppStorage();
      const container = document.createElement("div");
      container.id = "search-fixture";
      document.body.append(container);
      const details: import("../../src/tools/tool-details.ts").WebSearchDetails = { kind: "web_search", ok: false, provider: "jina", query: "fixture" };
      const descriptor = Object.getOwnPropertyDescriptor(storage, "settings");
      Object.defineProperty(storage, "settings", { configurable: true, get: () => { throw new Error("detection unavailable"); } });
      mountSearchSetupCard(container, details);
      if (descriptor) Object.defineProperty(storage, "settings", descriptor);
      setAppStorage(storage);
      mountSearchSetupCard(container, details);
    });
    await page.locator("#search-fixture .pi-search-setup__body").waitFor({ timeout: 3000 });
    assert.equal(await page.locator("#search-fixture .pi-search-setup").count(), 1);
  });
});

void test("code copy feedback resets when a copied block reconnects", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { CodeBlock } = await import("/src/ui/messages/code-block.ts") as typeof import("../../src/ui/messages/code-block.ts");
      Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: () => Promise.resolve() } });
      const block = new CodeBlock();
      block.id = "copy-fixture";
      block.code = "fixture";
      document.body.append(block);
    });
    await page.locator("#copy-fixture button").click();
    await page.locator("#copy-fixture button").filter({ hasText: "Copied" }).waitFor();
    await page.evaluate(() => {
      const block = document.getElementById("copy-fixture");
      if (!block) throw new Error("block missing");
      block.remove();
      document.body.append(block);
    });
    assert.equal(await page.locator("#copy-fixture button").innerText(), "Copy");
  });
});

async function setupConnection(page: Page): Promise<void> {
  await page.evaluate(async () => {
    const { getSettingsPagesDependencies } = await import("/src/commands/builtins/settings-pages/dependencies.ts") as typeof import("../../src/commands/builtins/settings-pages/dependencies.ts");
    const deps = getSettingsPagesDependencies().extensionsHub;
    if (!deps) throw new Error("Connections not initialized");
    deps.connectionManager.registerDefinition("fixture-owner", {
      id: "fixture.connection", title: "Lifecycle fixture", capability: "Fixture capability", authKind: "api_key",
      secretFields: [{ id: "key", label: "Fixture key", required: true }],
    });
    await deps.connectionManager.updateSecretsFromHost("fixture.connection", { key: "synthetic-fixture-key" });
  });
  await enterCommand(page, "/settings");
  await page.getByRole("button", { name: /^Connections/ }).click();
  const card = page.locator(".pi-item-card").filter({ hasText: "Lifecycle fixture" });
  await card.locator(".pi-item-card__header").click();
  await card.locator("input").waitFor({ state: "visible" });
}

void test("Connections commits only the latest complete refresh", async () => {
  await withTaskpane(async (page) => {
    await setupConnection(page);
    await page.evaluate(async () => {
      const { getSettingsPagesDependencies } = await import("/src/commands/builtins/settings-pages/dependencies.ts") as typeof import("../../src/commands/builtins/settings-pages/dependencies.ts");
      const manager = getSettingsPagesDependencies().extensionsHub?.connectionManager;
      if (!manager) throw new Error("manager missing");
      const get = manager.getSnapshot.bind(manager);
      let delay = true;
      manager.getSnapshot = async (id) => {
        const result = await get(id);
        if (id === "fixture.connection" && delay) {
          delay = false;
          document.body.dataset.refreshWaiting = "true";
          await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
          document.body.dataset.refreshReleased = "true";
        }
        return result;
      };
      const definition = manager.listDefinitions().find((entry) => entry.id === "fixture.connection");
      if (!definition) throw new Error("definition missing");
      manager.registerDefinition("fixture-owner", { ...definition, title: "Older fixture" });
    });
    await page.waitForFunction(() => document.body.dataset.refreshWaiting === "true");
    await page.evaluate(async () => {
      const { getSettingsPagesDependencies } = await import("/src/commands/builtins/settings-pages/dependencies.ts") as typeof import("../../src/commands/builtins/settings-pages/dependencies.ts");
      const manager = getSettingsPagesDependencies().extensionsHub?.connectionManager;
      const definition = manager?.listDefinitions().find((entry) => entry.id === "fixture.connection");
      if (!manager || !definition) throw new Error("definition missing");
      manager.registerDefinition("fixture-owner", { ...definition, title: "Newest fixture" });
    });
    await page.getByText("Newest fixture", { exact: true }).waitFor();
    await release(page);
    await page.waitForFunction(() => document.body.dataset.refreshReleased === "true");
    await page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())));
    assert.equal(await page.getByText("Older fixture", { exact: true }).count(), 0);
    assert.equal(await page.getByText("Newest fixture", { exact: true }).count(), 1);
  });
});

void test("Connections preserves a form focused after refresh starts", async () => {
  await withTaskpane(async (page) => {
    await setupConnection(page);
    await page.evaluate(async () => {
      const { getSettingsPagesDependencies } = await import("/src/commands/builtins/settings-pages/dependencies.ts") as typeof import("../../src/commands/builtins/settings-pages/dependencies.ts");
      const manager = getSettingsPagesDependencies().extensionsHub?.connectionManager;
      if (!manager) throw new Error("manager missing");
      const get = manager.getSnapshot.bind(manager);
      let delay = true;
      manager.getSnapshot = async (id) => {
        const result = await get(id);
        if (id === "fixture.connection" && delay) {
          delay = false;
          document.body.dataset.refreshWaiting = "true";
          await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
          document.body.dataset.refreshReleased = "true";
        }
        return result;
      };
      await manager.markValidated("fixture-owner", "fixture.connection");
    });
    await page.waitForFunction(() => document.body.dataset.refreshWaiting === "true");
    const input = page.locator(".pi-item-card").filter({ hasText: "Lifecycle fixture" }).locator("input");
    assert.equal(await input.count(), 1, "an in-flight refresh must preserve the current form");
    await input.fill("unsaved-fixture-draft");
    await release(page);
    await page.waitForFunction(() => document.body.dataset.refreshReleased === "true");
    await page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())));
    assert.equal(await input.inputValue(), "unsaved-fixture-draft");
    assert.equal(await input.evaluate((node) => node === document.activeElement), true);
  });
});

void test("extension secret saves share the page lock and retain newer input", async () => {
  await withTaskpane(async (page) => {
    await setupConnection(page);
    await page.evaluate(async () => {
      const { getSettingsPagesDependencies } = await import("/src/commands/builtins/settings-pages/dependencies.ts") as typeof import("../../src/commands/builtins/settings-pages/dependencies.ts");
      const manager = getSettingsPagesDependencies().extensionsHub?.connectionManager;
      if (!manager) throw new Error("manager missing");
      const update = manager.updateSecretsFromHost.bind(manager);
      const clear = manager.clearSecretsFromHost.bind(manager);
      manager.updateSecretsFromHost = async (...args) => {
        document.body.dataset.secretSaving = "true";
        await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
        await update(...args);
        document.body.dataset.secretSaved = "true";
      };
      manager.clearSecretsFromHost = async (...args) => { document.body.dataset.secretCleared = "true"; await clear(...args); };
    });
    const card = page.locator(".pi-item-card").filter({ hasText: "Lifecycle fixture" });
    await card.locator("input").fill("first-fixture-draft");
    await card.getByRole("button", { name: "Save", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.secretSaving === "true");
    await card.getByRole("button", { name: "Clear", exact: true }).click();
    assert.equal(await page.locator("body").getAttribute("data-secret-cleared"), null);
    await card.locator("input").fill("newer-fixture-draft");
    await release(page);
    await page.waitForFunction(() => document.body.dataset.secretSaved === "true");
    await page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())));
    assert.equal(await card.locator("input").inputValue(), "newer-fixture-draft");
  });
});

void test("Files ignores a workspace snapshot older than the last completed refresh", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/files");
    await page.locator(".pi-files-footer").waitFor();
    await page.evaluate(async () => {
      const { getFilesWorkspace } = await import("/src/files/workspace.ts") as typeof import("../../src/files/workspace.ts");
      const workspace = getFilesWorkspace();
      const list = workspace.listFiles.bind(workspace);
      let delay = true;
      workspace.listFiles = async (...args) => {
        const snapshot = await list(...args);
        if (delay) {
          delay = false;
          document.body.dataset.filesRefreshWaiting = "true";
          await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
          document.body.dataset.filesRefreshReleased = "true";
        }
        return snapshot;
      };
      await workspace.writeTextFile("first.txt", "first", "text/plain");
    });
    await page.waitForFunction(() => document.body.dataset.filesRefreshWaiting === "true");
    await page.evaluate(async () => {
      const { getFilesWorkspace } = await import("/src/files/workspace.ts") as typeof import("../../src/files/workspace.ts");
      await getFilesWorkspace().writeTextFile("second.txt", "second", "text/plain");
    });
    await page.getByText("second.txt", { exact: true }).waitFor();
    await release(page);
    await page.waitForFunction(() => document.body.dataset.filesRefreshReleased === "true");
    assert.equal(await page.getByText("second.txt", { exact: true }).count(), 1);
  });
});

void test("a cancelled Files action does not end another pending mutation", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getFilesWorkspace } = await import("/src/files/workspace.ts") as typeof import("../../src/files/workspace.ts");
      const { createFilesDialogDetailActions } = await import("/src/ui/files-dialog-actions.ts") as typeof import("../../src/ui/files-dialog-actions.ts");
      const workspace = getFilesWorkspace();
      await workspace.writeTextFile("pending.txt", "fixture", "text/plain");
      const file = (await workspace.listFiles()).find((entry) => entry.path === "pending.txt");
      if (!file) throw new Error("file missing");
      let pending = 0;
      const actions = createFilesDialogDetailActions({ file, fileRef: { path: file.path, locationKind: "workspace" }, auditContext: { actor: "user", source: "fixture" },
        workspace: { readFile: workspace.readFile.bind(workspace), downloadFile: workspace.downloadFile.bind(workspace), deleteFile: workspace.deleteFile.bind(workspace), renameFile: async (...args) => {
          await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
          await workspace.renameFile(...args);
        } },
        onMutationStateChange: (active) => { pending += active ? 1 : -1; document.body.dataset.pendingMutations = String(pending); },
        onAfterRename: () => Promise.resolve(), onAfterDelete: () => Promise.resolve(),
      });
      actions.id = "mutation-fixture";
      document.body.append(actions);
    });
    const actions = page.locator("#mutation-fixture");
    await actions.getByRole("button", { name: "Rename", exact: true }).click();
    const rename = page.locator("#pi-text-input-dialog-overlay");
    await rename.getByRole("textbox").fill("renamed.txt");
    await rename.getByRole("button", { name: "Rename", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.pendingMutations === "1");
    await actions.getByRole("button", { name: "Delete", exact: true }).click();
    await page.locator("#pi-confirm-dialog-overlay").getByRole("button", { name: "Cancel", exact: true }).last().click();
    assert.equal(await page.locator("body").getAttribute("data-pending-mutations"), "1");
    await release(page);
    await page.waitForFunction(() => document.body.dataset.pendingMutations === "0");
  });
});

void test("Rules publishes a newly identified workbook only after its existing rules load", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/rules");
    await page.getByPlaceholder(/Your preferences and habits/).waitFor();
    await page.evaluate(async () => {
      const { BrowserHost } = await import("/src/host/browser-host.ts") as typeof import("../../src/host/browser-host.ts");
      const { setCurrentSpreadsheetHost } = await import("/src/host/current.ts") as typeof import("../../src/host/current.ts");
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const settings = getAppStorage().settings;
      const get = settings.get.bind(settings);
      class IdentifiedHost extends BrowserHost {
        override getWorkbookContext(): Promise<import("../../src/workbook/context.ts").WorkbookContext> {
          return Promise.resolve({ workbookId: "identity-fixture", workbookName: "Fixture.xlsx", source: "document.url" });
        }
      }
      setCurrentSpreadsheetHost(new IdentifiedHost());
      settings.get = async (key) => {
        if (key === "workbook.instructions.v1.identity-fixture") {
          document.body.dataset.workbookRulesLoading = "true";
          await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
          return "Existing workbook rules";
        }
        return get(key);
      };
    });
    await page.getByRole("tab", { name: "This file", exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.workbookRulesLoading === "true");
    await page.getByRole("tab", { name: "All my files", exact: true }).click();
    await page.getByRole("tab", { name: "This file", exact: true }).click();
    assert.equal(await page.locator(".pi-overlay-textarea").isDisabled(), true);
    await release(page);
    await page.locator(".pi-overlay-textarea:not([disabled])").waitFor();
    assert.equal(await page.locator(".pi-overlay-textarea").inputValue(), "Existing workbook rules");
  });
});

void test("proxy URL save completion does not replace input typed during persistence", async () => {
  await withTaskpane(async (page) => {
    await enterCommand(page, "/settings");
    await page.getByRole("button", { name: /^Proxy/ }).click();
    const input = page.locator(".pi-settings-proxy-url");
    await input.waitFor();
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const settings = getAppStorage().settings;
      const set = settings.set.bind(settings);
      settings.set = async (...args) => {
        if (args[0] === "proxy.url") {
          document.body.dataset.urlSaving = "true";
          await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
        }
        await set(...args);
        document.body.dataset.urlSaved = "true";
      };
    });
    await input.fill("https://localhost:4180");
    await input.press("Enter");
    await page.waitForFunction(() => document.body.dataset.urlSaving === "true");
    await input.fill("https://localhost:4181");
    await release(page);
    await page.waitForFunction(() => document.body.dataset.urlSaved === "true");
    assert.equal(await input.inputValue(), "https://localhost:4181");
  });
});

void test("a paused duplicate toast stays paused without focus or hover transitions", async () => {
  await withTaskpane(async (page) => {
    await page.clock.install();
    await page.evaluate(async () => {
      const { showToast } = await import("/src/ui/toast.ts") as typeof import("../../src/ui/toast.ts");
      showToast("Hovered fixture", { id: "hovered-fixture", duration: 1000 });
    });
    const toast = page.locator(".pi-toast").filter({ hasText: "Hovered fixture" });
    await toast.hover();
    await page.evaluate(async () => {
      const { showToast } = await import("/src/ui/toast.ts") as typeof import("../../src/ui/toast.ts");
      showToast("Hovered fixture", { id: "hovered-fixture", duration: 1000 });
    });
    await page.clock.runFor(1100);
    assert.equal(await toast.count(), 1);
    await page.mouse.move(1, 1);
    await page.clock.runFor(1100);
    assert.equal(await toast.count(), 0);
  });
});

void test("proxy initial reads do not overwrite a URL edited before loading finishes", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { getAppStorage } = await import("/src/storage/local/app-storage.ts") as typeof import("../../src/storage/local/app-storage.ts");
      const settings = getAppStorage().settings;
      const get = settings.get.bind(settings);
      settings.get = async (key) => {
        if (key === "proxy.enabled") {
          document.body.dataset.proxyInitialRead = "true";
          await new Promise<void>((resolve) => document.addEventListener("fixture-release", () => resolve(), { once: true }));
        }
        const result = await get(key);
        if (key === "proxy.url") document.body.dataset.proxyInitialDone = "true";
        return result;
      };
    });
    await enterCommand(page, "/settings");
    await page.getByRole("button", { name: /^Proxy/ }).click();
    const input = page.locator(".pi-settings-proxy-url");
    await input.fill("https://localhost:4182");
    await release(page);
    await page.waitForFunction(() => document.body.dataset.proxyInitialDone === "true");
    assert.equal(await input.inputValue(), "https://localhost:4182");
  });
});

void test("Resume failures remain in the dialog and permit another attempt", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { showResumeDialog } = await import("/src/commands/builtins/resume-overlay.ts") as typeof import("../../src/commands/builtins/resume-overlay.ts");
      let attempt = 0;
      await showResumeDialog({ getRecentlyClosedItems: () => [{ id: "closed-fixture", sessionId: "closed-session", title: "Closed fixture", closedAt: new Date().toISOString(), workbookId: null }], onOpenSession: () => Promise.resolve(null), onReopenRecentlyClosed: () => {
        if (++attempt === 1) return Promise.reject(new Error("fixture reopen failed"));
        document.body.dataset.reopened = "true";
        return Promise.resolve(true);
      } });
    });
    await page.getByText("Closed fixture", { exact: true }).click();
    await page.getByText("Couldn't reopen session", { exact: true }).waitFor();
    await page.getByText("Closed fixture", { exact: true }).click();
    await page.waitForFunction(() => document.body.dataset.reopened === "true");
  });
});

void test("sidebar agent replacement clears the previous tool label", async () => {
  await withTaskpane(async (page) => {
    await page.evaluate(async () => {
      const { PiSidebar } = await import("/src/ui/pi-sidebar.ts") as typeof import("../../src/ui/pi-sidebar.ts");
      const sidebar = document.querySelector("pi-sidebar");
      if (!(sidebar instanceof PiSidebar) || !sidebar.agent) throw new Error("sidebar missing");
      const agent = sidebar.agent;
      let emit: (event: import("@earendil-works/pi-agent-core").AgentEvent) => void = () => {};
      agent.subscribe = (listener) => { emit = (event) => { void listener(event); }; return () => {}; };
      sidebar.agent = undefined;
      await sidebar.updateComplete;
      sidebar.agent = agent;
      await sidebar.updateComplete;
      agent.state.isStreaming = true;
      emit({ type: "tool_execution_start", toolCallId: "fixture-call", toolName: "fixture_old_tool", args: {} });
      await sidebar.updateComplete;
    });
    await page.getByRole("button", { name: "New tab", exact: true }).click();
    await page.locator(".pi-session-tab.is-active").filter({ hasText: "Chat 2" }).waitFor();
    await page.evaluate(async () => {
      const { PiSidebar } = await import("/src/ui/pi-sidebar.ts") as typeof import("../../src/ui/pi-sidebar.ts");
      const sidebar = document.querySelector("pi-sidebar");
      if (!(sidebar instanceof PiSidebar) || !sidebar.agent) throw new Error("sidebar missing");
      sidebar.agent.state.isStreaming = true;
      sidebar.syncFromAgent();
      await sidebar.updateComplete;
    });
    assert.notEqual(await page.locator(".pi-working__label").innerText(), "fixture_old_tool");
  });
});
