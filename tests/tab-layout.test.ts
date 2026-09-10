import assert from "node:assert/strict";
import { test } from "node:test";

import {
  loadWorkbookTabLayout,
  normalizeWorkbookTabLayout,
  parseWorkbookTabLayout,
  saveWorkbookTabLayout,
  workbookTabLayoutKey,
} from "../src/taskpane/tab-layout.ts";
import {
  ensureDefaultProxyUrl,
  readTaskpaneLanguage,
  readTaskpaneProxySettings,
  TASKPANE_LANGUAGE_SETTING_KEY,
  TASKPANE_PROXY_ENABLED_SETTING_KEY,
  TASKPANE_PROXY_URL_SETTING_KEY,
} from "../src/taskpane/settings.ts";

class MemoryTaskpaneSettingsStore {
  protected readonly values = new Map<string, unknown>();
  private readonly readError: Error | null;

  constructor(readError: Error | null = null) {
    this.readError = readError;
  }

  get(key: string): Promise<unknown> {
    if (this.readError) return Promise.reject(this.readError);
    return Promise.resolve(this.values.get(key) ?? null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }

  delete(key: string): Promise<void> {
    this.values.delete(key);
    return Promise.resolve();
  }
}

class ProxyEnabledReadFailureStore extends MemoryTaskpaneSettingsStore {
  override get(key: string): Promise<unknown> {
    if (key === TASKPANE_PROXY_ENABLED_SETTING_KEY) {
      return Promise.reject(new Error("proxy enabled read failed"));
    }
    return super.get(key);
  }
}

void test("workbookTabLayoutKey uses workbook id when available", () => {
  assert.equal(
    workbookTabLayoutKey("url_sha256:abc"),
    "workbook.tabLayout.v1.url_sha256:abc",
  );
});

void test("workbookTabLayoutKey falls back to global key when workbook id is missing", () => {
  assert.equal(workbookTabLayoutKey(null), "workbook.tabLayout.v1.__global__");
  assert.equal(workbookTabLayoutKey("   "), "workbook.tabLayout.v1.__global__");
});

void test("normalizeWorkbookTabLayout trims ids and falls back active tab", () => {
  const normalized = normalizeWorkbookTabLayout({
    sessionIds: ["  session-a  ", "", " session-b "],
    activeSessionId: "missing",
  });

  assert.deepEqual(normalized, {
    sessionIds: ["session-a", "session-b"],
    activeSessionId: "session-a",
  });
});

void test("parseWorkbookTabLayout accepts valid shape and normalizes active id", () => {
  const parsed = parseWorkbookTabLayout({
    sessionIds: [" session-a ", "session-b"],
    activeSessionId: "session-b",
  });

  assert.deepEqual(parsed, {
    sessionIds: ["session-a", "session-b"],
    activeSessionId: "session-b",
  });
});

void test("parseWorkbookTabLayout rejects invalid shapes", () => {
  assert.equal(parseWorkbookTabLayout(null), null);
  assert.equal(parseWorkbookTabLayout({}), null);
  assert.equal(parseWorkbookTabLayout({ sessionIds: "not-an-array" }), null);
  assert.equal(parseWorkbookTabLayout({ sessionIds: ["", "  "] }), null);
});

void test("tab layout reader round-trips valid data, defaults corrupt data, and propagates failed reads", async () => {
  const settings = new MemoryTaskpaneSettingsStore();
  const layout = {
    sessionIds: ["session-a", "session-b"],
    activeSessionId: "session-b",
  };

  await saveWorkbookTabLayout(settings, "workbook-a", layout);
  assert.deepEqual(await loadWorkbookTabLayout(settings, "workbook-a"), layout);
  assert.equal(await loadWorkbookTabLayout(settings, "missing"), null);

  await settings.set(workbookTabLayoutKey("corrupt"), { sessionIds: "not-an-array" });
  assert.equal(await loadWorkbookTabLayout(settings, "corrupt"), null);
  await assert.rejects(
    loadWorkbookTabLayout(new MemoryTaskpaneSettingsStore(new Error("read failed")), "workbook-a"),
    /read failed/u,
  );
});

void test("taskpane language reader accepts valid data and defaults on missing, corrupt, or failed reads", async () => {
  const settings = new MemoryTaskpaneSettingsStore();
  await settings.set(TASKPANE_LANGUAGE_SETTING_KEY, "zh-CN");
  assert.equal(await readTaskpaneLanguage(settings), "zh-CN");

  assert.equal(await readTaskpaneLanguage(new MemoryTaskpaneSettingsStore()), "en");
  await settings.set(TASKPANE_LANGUAGE_SETTING_KEY, { language: "zh-CN" });
  assert.equal(await readTaskpaneLanguage(settings), "en");
  assert.equal(
    await readTaskpaneLanguage(new MemoryTaskpaneSettingsStore(new Error("read failed"))),
    "en",
  );
});

void test("taskpane proxy reader round-trips current values and accepts legacy enabled flags", async () => {
  const settings = new MemoryTaskpaneSettingsStore();
  await settings.set(TASKPANE_PROXY_ENABLED_SETTING_KEY, true);
  await settings.set(TASKPANE_PROXY_URL_SETTING_KEY, " https://proxy.example.com/ ");
  assert.deepEqual(await readTaskpaneProxySettings(settings), {
    enabled: true,
    url: "https://proxy.example.com/",
  });

  for (const legacyEnabled of [1, "1", "true", "yes", "false", "0", "no", { enabled: false }]) {
    await settings.set(TASKPANE_PROXY_ENABLED_SETTING_KEY, legacyEnabled);
    assert.equal((await readTaskpaneProxySettings(settings)).enabled, true);
  }
});

void test("taskpane proxy reader defaults malformed values and propagates failed reads", async () => {
  assert.deepEqual(await readTaskpaneProxySettings(new MemoryTaskpaneSettingsStore()), {
    enabled: false,
    url: null,
  });

  const corrupt = new MemoryTaskpaneSettingsStore();
  await corrupt.set(TASKPANE_PROXY_ENABLED_SETTING_KEY, { enabled: true });
  await corrupt.set(TASKPANE_PROXY_URL_SETTING_KEY, { url: "https://proxy.example.com" });
  assert.deepEqual(await readTaskpaneProxySettings(corrupt), {
    enabled: true,
    url: null,
  });
  await assert.rejects(
    readTaskpaneProxySettings(new MemoryTaskpaneSettingsStore(new Error("read failed"))),
    /read failed/u,
  );
});

void test("taskpane startup never overwrites a custom proxy URL when another proxy read fails", async () => {
  const settings = new ProxyEnabledReadFailureStore();
  await settings.set(TASKPANE_PROXY_URL_SETTING_KEY, "https://custom.example.com");

  await ensureDefaultProxyUrl(settings, "office");

  assert.equal(
    await settings.get(TASKPANE_PROXY_URL_SETTING_KEY),
    "https://custom.example.com",
  );
});

void test("taskpane startup replaces a malformed proxy URL with the runtime default", async () => {
  const settings = new MemoryTaskpaneSettingsStore();
  await settings.set(TASKPANE_PROXY_ENABLED_SETTING_KEY, false);
  await settings.set(TASKPANE_PROXY_URL_SETTING_KEY, { url: "https://invalid.example.com" });

  await ensureDefaultProxyUrl(settings, "office");

  assert.equal(await settings.get(TASKPANE_PROXY_URL_SETTING_KEY), "https://localhost:3003");
});
