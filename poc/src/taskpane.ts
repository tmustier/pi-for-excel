/**
 * Pi for Excel — PoC Test Harness
 *
 * Tests platform capabilities in the Office.js webview.
 * This is throw-away code. Don't judge it.
 */

// Pre-load pi-web-ui CSS so it's ready when ChatPanel mounts
import "@mariozechner/pi-web-ui/app.css";

// ==========================================================================
// FIX: Lit class field shadowing (tsgo bug)
//
// pi-web-ui is compiled with tsgo which emits native class field declarations
// despite useDefineForClassFields:false. Native class fields use [[Define]]
// semantics, creating own properties that shadow Lit's @state() prototype
// accessors. Lit's dev-mode check in performUpdate() throws on this.
//
// Fix: monkey-patch ReactiveElement.prototype.performUpdate to auto-fix
// shadowed properties before the check runs. This affects ALL Lit components.
// See: https://lit.dev/msg/class-field-shadowing
// ==========================================================================
import { ReactiveElement } from "lit";

const _origPerformUpdate = ReactiveElement.prototype.performUpdate;
ReactiveElement.prototype.performUpdate = function (this: ReactiveElement) {
  if (!this.hasUpdated) {
    // Fix ALL own properties that shadow prototype accessors (get/set).
    // This handles @state(), @property(), @query(), @queryAll(), etc.
    const proto = Object.getPrototypeOf(this);
    for (const key of Object.getOwnPropertyNames(this)) {
      // Skip known LitElement internal properties
      if (key.startsWith("__") || key === "renderRoot" || key === "isUpdatePending" || key === "hasUpdated") continue;
      const protoDesc = Object.getOwnPropertyDescriptor(proto, key);
      if (protoDesc && (protoDesc.get || protoDesc.set)) {
        // Own data property shadows a prototype accessor — fix it
        const ownDesc = Object.getOwnPropertyDescriptor(this, key);
        if (ownDesc && !ownDesc.get && !ownDesc.set) {
          const value = (this as any)[key];
          delete (this as any)[key];
          // Only re-assign if there's a setter (skip @query getters with no setter)
          if (protoDesc.set) {
            (this as any)[key] = value;
          }
        }
      }
    }
  }
  return _origPerformUpdate.call(this);
};


// ============================================================================
// LOGGING
// ============================================================================

const logEl = document.getElementById("log")!;
const statusEl = document.getElementById("status")!;

function log(msg: string, type: "ok" | "err" | "info" = "info") {
  const cls = type === "ok" ? "log-ok" : type === "err" ? "log-err" : "log-info";
  const prefix = type === "ok" ? "✅" : type === "err" ? "❌" : "ℹ️";
  logEl.innerHTML += `<span class="${cls}">${prefix} ${msg}</span>\n`;
  logEl.scrollTop = logEl.scrollHeight;
  console.log(`[${type}]`, msg);
}

function markButton(id: string, pass: boolean) {
  const btn = document.getElementById(id);
  if (btn) btn.className = pass ? "pass" : "fail";
}

// ============================================================================
// OFFICE.JS INITIALIZATION
// ============================================================================

let officeReady = false;

// @ts-ignore - Office is loaded via CDN script tag
Office.onReady((info: { host: any; platform: any }) => {
  officeReady = true;
  statusEl.className = "ready";
  statusEl.textContent = `Office.js ready — Host: ${info.host}, Platform: ${info.platform}`;
  log(`Office.js initialized: host=${info.host}, platform=${info.platform}`, "ok");
});

// Fallback: if not in Excel, still let webview tests work
setTimeout(() => {
  if (!officeReady) {
    statusEl.className = "error";
    statusEl.textContent = "Office.js not available (running outside Excel?)";
    log("Office.js not available — Excel tests will fail, webview tests will still work", "err");
  }
}, 3000);

// ============================================================================
// 1. OFFICE.JS BASICS
// ============================================================================

async function testReadCells() {
  const id = "btn-read";
  try {
    // @ts-ignore
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:C3");
      range.load("values,formulas,address,numberFormat");
      await context.sync();

      log(`Read ${range.address}:`, "ok");
      log(`  Values: ${JSON.stringify(range.values)}`, "info");
      log(`  Formulas: ${JSON.stringify(range.formulas)}`, "info");
      log(`  Number formats: ${JSON.stringify(range.numberFormat)}`, "info");
    });
    markButton(id, true);
  } catch (e: any) {
    log(`Read failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

async function testWriteCells() {
  const id = "btn-write";
  try {
    // @ts-ignore
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("E1:G3");
      range.values = [
        ["Pi PoC", "Test", new Date().toLocaleTimeString()],
        [42, "=E1&\" works!\"", "=SUM(E2,100)"],
        ["=1/0", "=E2+1", "done"],
      ];
      range.format.autofitColumns();
      await context.sync();

      // Read back to verify
      const verify = sheet.getRange("E1:G3");
      verify.load("values,formulas");
      await context.sync();
      log(`Write + readback:`, "ok");
      log(`  Values: ${JSON.stringify(verify.values)}`, "info");
      log(`  Formulas: ${JSON.stringify(verify.formulas)}`, "info");

      // Check for errors in values
      const flat = verify.values.flat();
      const errors = flat.filter(
        (v: any) => typeof v === "string" && v.startsWith("#")
      );
      if (errors.length > 0) {
        log(`  Formula errors detected: ${JSON.stringify(errors)}`, "err");
      }
    });
    markButton(id, true);
  } catch (e: any) {
    log(`Write failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

async function testOverview() {
  const id = "btn-overview";
  try {
    // @ts-ignore
    await Excel.run(async (context: any) => {
      const wb = context.workbook;
      wb.load("name");
      const sheets = wb.worksheets;
      sheets.load("items/name,items/id,items/position");
      await context.sync();

      log(`Workbook: ${wb.name}`, "ok");
      log(`Sheets (${sheets.items.length}):`, "info");

      for (const sheet of sheets.items) {
        const used = sheet.getUsedRange();
        used.load("rowCount,columnCount,address");
        // Try to get first row as headers
        const headerRange = sheet.getRange("1:1").getUsedRangeOrNullObject();
        headerRange.load("values,address");
        await context.sync();

        const dims = used.isNullObject
          ? "empty"
          : `${used.rowCount}×${used.columnCount}`;
        const headers = headerRange.isNullObject
          ? []
          : headerRange.values[0].filter((v: any) => v !== "");
        log(
          `  ${sheet.position + 1}. "${sheet.name}" (${dims}) — headers: [${headers.slice(0, 5).join(", ")}${headers.length > 5 ? "..." : ""}]`,
          "info"
        );
      }
    });
    markButton(id, true);
  } catch (e: any) {
    log(`Overview failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

async function testSelection() {
  const id = "btn-selection";
  try {
    // @ts-ignore
    await Excel.run(async (context: any) => {
      const range = context.workbook.getSelectedRange();
      range.load("address,values,formulas,worksheet/name");
      await context.sync();

      log(`Selection: ${range.worksheet.name}!${range.address}`, "ok");
      log(`  Values: ${JSON.stringify(range.values)}`, "info");
      if (range.formulas.flat().some((f: any) => f !== "" && f !== range.values.flat()[range.formulas.flat().indexOf(f)])) {
        log(`  Formulas: ${JSON.stringify(range.formulas)}`, "info");
      }
    });
    markButton(id, true);
  } catch (e: any) {
    log(`Selection failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

// ============================================================================
// 2. ADVANCED OFFICE.JS
// ============================================================================

async function testPrecedents() {
  const id = "btn-precedents";
  try {
    // @ts-ignore
    await Excel.run(async (context: any) => {
      const range = context.workbook.getSelectedRange();
      range.load("address,formulas");
      await context.sync();

      log(`Testing getDirectPrecedents on ${range.address}`, "info");
      log(`  Formula: ${JSON.stringify(range.formulas)}`, "info");

      // This is a preview API — may not work
      const precedents = range.getDirectPrecedents();
      precedents.load("addresses");
      await context.sync();

      log(`Direct precedents:`, "ok");
      for (const area of precedents.addresses) {
        log(`  ${JSON.stringify(area)}`, "info");
      }
    });
    markButton(id, true);
  } catch (e: any) {
    log(`getDirectPrecedents: ${e.message}`, "err");
    if (e.message.includes("not supported") || e.message.includes("not a function")) {
      log(`  → API not available in this environment. Fallback: parse formula strings.`, "info");
    }
    markButton(id, false);
  }
}

async function testEvents() {
  const id = "btn-events";
  try {
    // @ts-ignore
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      sheet.onChanged.add((event: any) => {
        log(
          `  🔔 Change detected: ${event.address} on "${event.worksheetId}" — type: ${event.changeType}`,
          "ok"
        );
      });
      await context.sync();

      log(`Event listener registered on "${sheet.name}". Edit any cell to test.`, "ok");
      log(`  (Listener active for this session — edit a cell now!)`, "info");
    });
    markButton(id, true);
  } catch (e: any) {
    log(`Events failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

async function testNamedRanges() {
  const id = "btn-named";
  try {
    // @ts-ignore
    await Excel.run(async (context: any) => {
      const names = context.workbook.names;
      names.load("items/name,items/type,items/value,items/visible");
      await context.sync();

      if (names.items.length === 0) {
        log(`No named ranges found in workbook`, "info");
        // Create one to test
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        context.workbook.names.add("PiTestRange", sheet.getRange("A1"), "PoC test");
        await context.sync();
        log(`  Created test named range "PiTestRange" → A1`, "ok");

        // Clean up
        const created = context.workbook.names.getItem("PiTestRange");
        created.delete();
        await context.sync();
        log(`  Deleted test named range`, "info");
      } else {
        log(`Named ranges (${names.items.length}):`, "ok");
        for (const n of names.items) {
          log(`  ${n.name} = ${n.value} (${n.type}, visible: ${n.visible})`, "info");
        }
      }
    });
    markButton(id, true);
  } catch (e: any) {
    log(`Named ranges failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

async function testConditionalFormat() {
  const id = "btn-condfmt";
  try {
    // @ts-ignore
    await Excel.run(async (context: any) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:A10");
      const cfs = range.conditionalFormats;
      cfs.load("items/type,items/id");
      await context.sync();

      log(`Conditional formats on A1:A10: ${cfs.items.length} rules`, "ok");

      // Try adding one
      const cf = cfs.add(
        // @ts-ignore
        Excel.ConditionalFormatType.cellValue
      );
      cf.cellValue.format.font.color = "#FF0000";
      cf.cellValue.rule = {
        formula1: "=0",
        // @ts-ignore
        operator: Excel.ConditionalCellValueOperator.greaterThan,
      };
      await context.sync();
      log(`  Added test conditional format (values > 0 → red)`, "ok");

      // Remove it
      cf.delete();
      await context.sync();
      log(`  Removed test conditional format`, "info");
    });
    markButton(id, true);
  } catch (e: any) {
    log(`Conditional formatting failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

// ============================================================================
// 3. WEBVIEW CAPABILITIES
// ============================================================================

async function testIndexedDB() {
  const id = "btn-indexeddb";
  try {
    const request = indexedDB.open("pi-poc-test", 1);
    await new Promise<void>((resolve, reject) => {
      request.onupgradeneeded = () => {
        request.result.createObjectStore("test");
      };
      request.onsuccess = () => {
        const db = request.result;
        const tx = db.transaction("test", "readwrite");
        const store = tx.objectStore("test");
        store.put({ timestamp: Date.now(), data: "hello from PoC" }, "test-key");
        tx.oncomplete = () => {
          // Read it back
          const tx2 = db.transaction("test", "readonly");
          const store2 = tx2.objectStore("test");
          const get = store2.get("test-key");
          get.onsuccess = () => {
            log(`IndexedDB write + read: ${JSON.stringify(get.result)}`, "ok");
            db.close();
            // Clean up
            indexedDB.deleteDatabase("pi-poc-test");
            resolve();
          };
          get.onerror = () => reject(get.error);
        };
      };
      request.onerror = () => reject(request.error);
    });
    markButton(id, true);
  } catch (e: any) {
    log(`IndexedDB failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

async function testCORS() {
  const id = "btn-cors";
  log(`Testing CORS to various endpoints...`, "info");

  // Test 1: Generic HTTPS fetch
  try {
    const r = await fetch("https://httpbin.org/get", { method: "GET" });
    log(`  httpbin.org: ${r.status} ${r.ok ? "✓" : "✗"}`, r.ok ? "ok" : "err");
  } catch (e: any) {
    log(`  httpbin.org: BLOCKED — ${e.message}`, "err");
  }

  // Test 2: OpenAI-compatible endpoint (just test CORS headers, don't need key)
  try {
    const r = await fetch("https://api.openai.com/v1/models", {
      method: "GET",
      headers: { Authorization: "Bearer test-not-a-real-key" },
    });
    log(`  api.openai.com: ${r.status} (CORS ${r.status === 0 ? "blocked" : "allowed"})`, r.status !== 0 ? "ok" : "err");
  } catch (e: any) {
    log(`  api.openai.com: ${e.message}`, "err");
  }

  // Test 3: Anthropic (requires special header for browser access)
  try {
    const r = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": "test-not-a-real-key",
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 1, messages: [{ role: "user", content: "hi" }] }),
    });
    log(`  api.anthropic.com: ${r.status} (CORS allowed with dangerous-direct-browser-access header)`, r.status !== 0 ? "ok" : "err");
  } catch (e: any) {
    log(`  api.anthropic.com (with header): ${e.message}`, "err");
    // Try without the header to confirm it's the issue
    try {
      const r2 = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": "test-not-a-real-key",
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 1, messages: [{ role: "user", content: "hi" }] }),
      });
      log(`  api.anthropic.com (without header): ${r2.status}`, "info");
    } catch {
      log(`  api.anthropic.com (without header): also blocked — confirms header is required`, "info");
    }
  }

  // Test 4: Google AI
  try {
    const r = await fetch("https://generativelanguage.googleapis.com/v1beta/models?key=test-not-a-key");
    log(`  googleapis.com: ${r.status} (CORS ${r.status === 0 ? "blocked" : "allowed"})`, r.status !== 0 ? "ok" : "err");
  } catch (e: any) {
    log(`  googleapis.com: ${e.message}`, "err");
  }

  markButton(id, true); // Always mark as done, individual results shown
}

async function testWASM() {
  const id = "btn-wasm";
  try {
    // Minimal WASM test: compile and run a simple addition function
    const bytes = new Uint8Array([
      0x00, 0x61, 0x73, 0x6d, // magic
      0x01, 0x00, 0x00, 0x00, // version
      0x01, 0x07, 0x01, 0x60, 0x02, 0x7f, 0x7f, 0x01, 0x7f, // type: (i32, i32) -> i32
      0x03, 0x02, 0x01, 0x00, // function
      0x07, 0x07, 0x01, 0x03, 0x61, 0x64, 0x64, 0x00, 0x00, // export "add"
      0x0a, 0x09, 0x01, 0x07, 0x00, 0x20, 0x00, 0x20, 0x01, 0x6a, 0x0b, // code: local.get 0 + local.get 1 + i32.add
    ]);

    const module = await WebAssembly.compile(bytes);
    const instance = await WebAssembly.instantiate(module);
    const add = (instance.exports as any).add as (a: number, b: number) => number;
    const result = add(40, 2);

    log(`WASM: add(40, 2) = ${result} ${result === 42 ? "✓" : "✗"}`, result === 42 ? "ok" : "err");
    markButton(id, result === 42);
  } catch (e: any) {
    log(`WASM failed: ${e.message}`, "err");
    markButton(id, false);
  }
}

async function testPyodide() {
  const id = "btn-pyodide";
  log(`Loading Pyodide (this may take 10-20s on first load)...`, "info");
  try {
    // Dynamic import — Pyodide is heavy, only load when tested
    const { loadPyodide } = await import("pyodide");
    const pyodide = await loadPyodide({
      indexURL: "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/",
    });

    log(`Pyodide loaded: Python ${pyodide.version}`, "ok");

    // Test 1: Basic Python
    const result1 = pyodide.runPython(`
import sys
f"Python {sys.version}"
    `);
    log(`  Python version: ${result1}`, "info");

    // Test 2: Can we import numpy/pandas?
    await pyodide.loadPackage("numpy");
    const result2 = pyodide.runPython(`
import numpy as np
arr = np.array([1, 2, 3, 4, 5])
f"numpy: mean={arr.mean()}, std={arr.std():.4f}"
    `);
    log(`  ${result2}`, "ok");

    // Test 3: pandas
    await pyodide.loadPackage("pandas");
    const result3 = pyodide.runPython(`
import pandas as pd
df = pd.DataFrame({"Revenue": [100, 120, 145], "Cost": [80, 90, 105]})
df["Profit"] = df["Revenue"] - df["Cost"]
df.to_string()
    `);
    log(`  pandas:\n${result3}`, "ok");

    // Test 4: Pass data from JS to Python and back
    const testData = [[1, 2, 3], [4, 5, 6], [7, 8, 9]];
    pyodide.globals.set("js_data", pyodide.toPy(testData));
    const result4 = pyodide.runPython(`
import json
# js_data is already a Python list (toPy converted it)
data = js_data
total = sum(sum(row) for row in data)
json.dumps({"total": total, "shape": f"{len(data)}x{len(data[0])}"})
    `);
    log(`  JS→Python→JS roundtrip: ${result4}`, "ok");

    markButton(id, true);
  } catch (e: any) {
    log(`Pyodide failed: ${e.message}`, "err");
    log(`  Stack: ${e.stack?.split("\n").slice(0, 3).join("\n")}`, "info");
    markButton(id, false);
  }
}

// ============================================================================
// 4. PI INTEGRATION
// ============================================================================

async function testChatPanel() {
  const id = "btn-chat";
  const container = document.getElementById("chat-container")!;

  // Capture ALL errors (including from Lit's async lifecycle)
  const prevOnError = window.onerror;
  const errors: string[] = [];
  window.onerror = (msg, source, lineno, colno, error) => {
    errors.push(`${msg} at ${source}:${lineno}:${colno}`);
    log(`  ⚠️ Global error: ${msg}`, "err");
    if (prevOnError) prevOnError(msg, source!, lineno!, colno!, error!);
    return false;
  };
  window.addEventListener("unhandledrejection", (e) => {
    errors.push(`Unhandled rejection: ${e.reason}`);
    log(`  ⚠️ Unhandled rejection: ${e.reason}`, "err");
  });

  try {
    log(`Importing pi-web-ui...`, "info");

    const { Agent } = await import("@mariozechner/pi-agent-core");
    const { getModel } = await import("@mariozechner/pi-ai");
    const {
      ChatPanel,
      AppStorage,
      IndexedDBStorageBackend,
      ProviderKeysStore,
      SessionsStore,
      SettingsStore,
      setAppStorage,
      ApiKeyPromptDialog,
      SettingsDialog,
      ProvidersModelsTab,
    } = await import("@mariozechner/pi-web-ui");
    const { html, render } = await import("lit");

    log(`pi-web-ui imported successfully`, "ok");

    // Set up storage (same as pi-web-ui example)
    const settings = new SettingsStore();
    const providerKeys = new ProviderKeysStore();
    const sessions = new SessionsStore();

    const backend = new IndexedDBStorageBackend({
      dbName: "pi-excel-poc",
      version: 1,
      stores: [
        settings.getConfig(),
        providerKeys.getConfig(),
        sessions.getConfig(),
        SessionsStore.getMetadataConfig(),
      ],
    });

    settings.setBackend(backend);
    providerKeys.setBackend(backend);
    sessions.setBackend(backend);

    const storage = new AppStorage(settings, providerKeys, sessions, undefined, backend);
    setAppStorage(storage);

    // Create agent
    const agent = new Agent({
      initialState: {
        systemPrompt: `You are Pi, an AI assistant embedded in Microsoft Excel. You can read and write spreadsheet data. This is a proof-of-concept test.`,
        model: getModel("anthropic", "claude-sonnet-4-5-20250929"),
        thinkingLevel: "off",
        messages: [],
        tools: [],
      },
    });

    // Create ChatPanel (same pattern as pi-web-ui example)
    const chatPanel = new ChatPanel();

    await chatPanel.setAgent(agent, {
      onApiKeyRequired: async (provider: string) => {
        return await ApiKeyPromptDialog.prompt(provider);
      },
    });

    container.classList.add("visible");

    // Use lit's render() — the EXACT pattern from pi-web-ui example
    render(html`
      <div class="w-full h-full flex flex-col bg-background text-foreground overflow-hidden">
        ${chatPanel}
      </div>
    `, container);

    // Show the Settings button
    document.getElementById("btn-settings")!.style.display = "";

    // Expose openSettings globally
    (window as any).openSettings = () => {
      SettingsDialog.open([new ProvidersModelsTab()]);
    };

    // Expose OAuth login globally
    (window as any).oauthLogin = async (providerId: string) => {
      await doOAuthLogin(providerId, providerKeys);
    };
    document.getElementById("btn-oauth")!.style.display = "";

    // Auto-restore auth from pi's auth.json (dev) or localStorage (browser OAuth)
    await restoreAuthCredentials(providerKeys);

    log(`ChatPanel mounted via lit render()! Try chatting.`, "ok");
    log(`  Click 🔑 Login to authenticate with your provider subscription (free).`, "info");
    log(`  Click ⚙️ Settings to add API keys instead.`, "info");

    // Debug after render completes
    await chatPanel.updateComplete;
    log(`  updateComplete resolved. innerHTML length=${chatPanel.innerHTML.length}`, "info");

    // Additional debug after a delay
    setTimeout(() => {
      log(`  DOM debug (2s later):`, "info");
      log(`    chatPanel.innerHTML length=${chatPanel.innerHTML.length}`, "info");
      log(`    chatPanel.children=${chatPanel.children.length}`, "info");
      const ai = chatPanel.querySelector("agent-interface");
      log(`    agent-interface found=${!!ai}`, "info");
      if (ai) {
        log(`    agent-interface.innerHTML length=${ai.innerHTML.length}`, "info");
        const ta = ai.querySelector("textarea");
        log(`    textarea found=${!!ta}`, "info");
        if (ta) {
          log(`    textarea visible=${ta.offsetHeight > 0}, height=${ta.offsetHeight}`, "info");
        }
      }

      // If still empty, try to diagnose the Lit lifecycle
      if (chatPanel.innerHTML.length === 0) {
        log(`  ❌ ChatPanel rendered NOTHING. Diagnosing...`, "err");
        log(`    chatPanel.isConnected=${chatPanel.isConnected}`, "info");
        log(`    chatPanel.agent set=${!!chatPanel.agent}`, "info");
        log(`    chatPanel.agentInterface set=${!!chatPanel.agentInterface}`, "info");

        // Check if own properties shadow prototype accessors
        const proto = Object.getPrototypeOf(chatPanel);
        const agentDesc = Object.getOwnPropertyDescriptor(chatPanel, "agent");
        const protoAgentDesc = Object.getOwnPropertyDescriptor(proto, "agent");
        log(`    Own 'agent' descriptor: ${JSON.stringify(agentDesc ? { value: typeof agentDesc.value, writable: agentDesc.writable, get: !!agentDesc.get, set: !!agentDesc.set } : null)}`, "info");
        log(`    Prototype 'agent' descriptor: ${JSON.stringify(protoAgentDesc ? { get: !!protoAgentDesc.get, set: !!protoAgentDesc.set } : null)}`, "info");

        // Try manual render
        log(`  Trying manual render...`, "info");
        try {
          const result = (chatPanel as any).render();
          log(`    render() returned: ${result ? "TemplateResult" : "falsy"} (type: ${typeof result})`, "info");
          if (result) {
            // Stamp it manually using lit's render
            render(result, chatPanel);
            log(`    Manual render stamped. innerHTML length=${chatPanel.innerHTML.length}`, result ? "ok" : "err");
          }
        } catch (e: any) {
          log(`    Manual render failed: ${e.message}`, "err");
        }
      }

      if (errors.length > 0) {
        log(`  Captured ${errors.length} error(s) during lifecycle`, "err");
      }
    }, 2000);

    markButton(id, true);
  } catch (e: any) {
    log(`ChatPanel failed: ${e.message}`, "err");
    log(`  Stack: ${e.stack?.split("\n").slice(0, 5).join("\n")}`, "info");
    markButton(id, false);
  }
}

// ============================================================================
// OAUTH CORS PROXY
// Route OAuth token exchange requests through Vite's dev server proxy
// (these endpoints don't support browser CORS)
// ============================================================================
const OAUTH_PROXY_REWRITES: [string, string][] = [
  // OAuth token endpoints
  ["https://console.anthropic.com/", "/oauth-proxy/anthropic/"],
  ["https://github.com/", "/oauth-proxy/github/"],
  ["https://auth.openai.com/", "/api-proxy/openai-auth/"],
  ["https://oauth2.googleapis.com/", "/api-proxy/google-oauth/"],
  // API endpoints
  ["https://api.anthropic.com/", "/api-proxy/anthropic/"],
  ["https://api.openai.com/", "/api-proxy/openai/"],
  ["https://chatgpt.com/", "/api-proxy/chatgpt/"],
  ["https://generativelanguage.googleapis.com/", "/api-proxy/google/"],
];

const _originalFetch = window.fetch.bind(window);
window.fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
  let url = typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
  let proxied = false;
  for (const [prefix, proxy] of OAUTH_PROXY_REWRITES) {
    if (url.startsWith(prefix)) {
      console.log(`[fetch-proxy] Rewriting ${url.substring(0, 60)}... → ${proxy}...`);
      url = url.replace(prefix, proxy);
      proxied = true;
      break;
    }
  }

  if (proxied) {
    // Strip the "anthropic-dangerous-direct-browser-access" header when proxying.
    // The proxy makes it a server-to-server request, but Anthropic checks this
    // header to apply org CORS policy — which blocks OAuth tokens.
    const headers = new Headers(init?.headers);
    headers.delete("anthropic-dangerous-direct-browser-access");

    // Also strip headers that mark this as a browser request
    const newInit = { ...init, headers };

    if (typeof input !== "string" && !(input instanceof URL) && input instanceof Request) {
      // Clone the request with new URL and headers
      const newHeaders = new Headers(input.headers);
      newHeaders.delete("anthropic-dangerous-direct-browser-access");
      input = new Request(url, { ...input, headers: newHeaders });
    } else {
      input = url;
    }
    return _originalFetch(input, newInit);
  }

  return _originalFetch(input, init);
};

// ============================================================================
// OAUTH LOGIN (browser-compatible flows)
// ============================================================================

// Providers whose OAuth flows are 100% browser-compatible
// (no node:http/node:crypto — just fetch + Web Crypto)
const BROWSER_OAUTH_PROVIDERS = ["anthropic", "github-copilot"];

/** Simple modal helper — returns user input or null if cancelled */
function showModal(title: string, bodyHtml: string): Promise<string | null> {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;";
    const modal = document.createElement("div");
    modal.style.cssText = "background:#fff;border-radius:12px;padding:24px;max-width:480px;width:90%;box-shadow:0 8px 32px rgba(0,0,0,0.2);font-family:system-ui;";
    modal.innerHTML = `
      <h3 style="margin:0 0 16px;font-size:18px;">${title}</h3>
      <div>${bodyHtml}</div>
      <div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;">
        <button id="modal-cancel" style="padding:8px 16px;border:1px solid #ddd;border-radius:6px;background:#fff;cursor:pointer;">Cancel</button>
      </div>
    `;
    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    const cleanup = (result: string | null) => {
      document.body.removeChild(overlay);
      resolve(result);
    };

    overlay.addEventListener("click", (e) => { if (e.target === overlay) cleanup(null); });
    modal.querySelector("#modal-cancel")!.addEventListener("click", () => cleanup(null));
  });
}

function showInputModal(title: string, message: string, placeholder?: string): Promise<string | null> {
  return new Promise((resolve) => {
    const overlay = document.createElement("div");
    overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;";
    const modal = document.createElement("div");
    modal.style.cssText = "background:#fff;border-radius:12px;padding:24px;max-width:480px;width:90%;box-shadow:0 8px 32px rgba(0,0,0,0.2);font-family:system-ui;";
    modal.innerHTML = `
      <h3 style="margin:0 0 12px;font-size:18px;">${title}</h3>
      <p style="margin:0 0 12px;color:#555;font-size:14px;">${message}</p>
      <input id="modal-input" type="text" placeholder="${placeholder || ''}" 
        style="width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px;box-sizing:border-box;" />
      <div style="margin-top:16px;display:flex;gap:8px;justify-content:flex-end;">
        <button id="modal-cancel" style="padding:8px 16px;border:1px solid #ddd;border-radius:6px;background:#fff;cursor:pointer;">Cancel</button>
        <button id="modal-ok" style="padding:8px 16px;border:none;border-radius:6px;background:#2563eb;color:#fff;cursor:pointer;">Submit</button>
      </div>
    `;
    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    const input = modal.querySelector("#modal-input") as HTMLInputElement;
    input.focus();

    const cleanup = (result: string | null) => {
      document.body.removeChild(overlay);
      resolve(result);
    };

    overlay.addEventListener("click", (e) => { if (e.target === overlay) cleanup(null); });
    modal.querySelector("#modal-cancel")!.addEventListener("click", () => cleanup(null));
    modal.querySelector("#modal-ok")!.addEventListener("click", () => cleanup(input.value));
    input.addEventListener("keydown", (e) => { if (e.key === "Enter") cleanup(input.value); if (e.key === "Escape") cleanup(null); });
  });
}

/**
 * Load auth credentials from pi's ~/.pi/agent/auth.json (served by Vite plugin)
 * and from localStorage (for browser-originated OAuth sessions).
 * This means any provider already authenticated in pi TUI works automatically.
 */
async function restoreAuthCredentials(providerKeys: any) {
  const { getOAuthProvider } = await import("@mariozechner/pi-ai");

  // Map OAuth provider IDs to the API provider names used by pi-ai.
  // IMPORTANT: openai-codex stays as "openai-codex" — it uses chatgpt.com/backend-api,
  // NOT api.openai.com. pi-ai has separate models/provider for it.
  const providerMap: Record<string, string> = {
    "anthropic": "anthropic",
    "openai-codex": "openai-codex",
    "github-copilot": "github-copilot",
    "gemini-cli": "google",
    "google-gemini-cli": "google",
    "antigravity": "google",
    "google-antigravity": "google",
  };

  // 1. Try loading from pi's auth.json (dev server only)
  try {
    const res = await _originalFetch("/__pi-auth");
    if (res.ok) {
      const authData = await res.json() as Record<string, any>;
      log(`  Found pi auth.json with ${Object.keys(authData).length} provider(s)`, "info");

      for (const [providerId, cred] of Object.entries(authData)) {
        try {
          if (cred.type === "api_key" && cred.key) {
            // Direct API key — use as-is
            const apiProvider = providerMap[providerId] || providerId;
            await providerKeys.set(apiProvider, cred.key);
            log(`  ✅ ${providerId}: API key loaded from pi`, "ok");

          } else if (cred.type === "oauth") {
            const provider = getOAuthProvider(providerId);
            if (!provider) {
              log(`  ⚠️ ${providerId}: no OAuth provider registered, skipping`, "info");
              continue;
            }

            const apiProvider = providerMap[providerId] || providerId;

            if (Date.now() >= cred.expires) {
              // Token expired — try to refresh through proxy
              log(`  ${providerId}: token expired, refreshing...`, "info");
              try {
                const refreshed = await provider.refreshToken(cred);
                const apiKey = provider.getApiKey(refreshed);
                await providerKeys.set(apiProvider, apiKey);
                log(`  ✅ ${providerId}: token refreshed`, "ok");
              } catch (e: any) {
                log(`  ⚠️ ${providerId}: refresh failed (${e.message})`, "err");
              }
            } else {
              const apiKey = provider.getApiKey(cred);
              await providerKeys.set(apiProvider, apiKey);
              const hours = Math.round((cred.expires - Date.now()) / 3600000);
              log(`  ✅ ${providerId}: OAuth token loaded (expires in ${hours}h)`, "ok");
            }
          }
        } catch (e: any) {
          log(`  ⚠️ ${providerId}: failed to load (${e.message})`, "err");
        }
      }
      // Clear any stale keys from old provider mappings
      // (e.g., we previously mapped openai-codex → openai incorrectly)
      const allStored = await providerKeys.getAll?.() || {};
      for (const staleProvider of Object.keys(allStored)) {
        if (!Object.values(providerMap).includes(staleProvider) && !authData[staleProvider]) {
          log(`  🧹 Clearing stale key for "${staleProvider}"`, "info");
          await providerKeys.set(staleProvider, "");
        }
      }

      return; // pi auth loaded, skip localStorage fallback
    }
  } catch {
    // Not running with Vite dev server — fall through to localStorage
  }

  // 2. Fallback: restore from localStorage (browser OAuth sessions)
  for (const providerId of BROWSER_OAUTH_PROVIDERS) {
    const stored = localStorage.getItem(`oauth_${providerId}`);
    if (!stored) continue;

    try {
      const credentials = JSON.parse(stored);
      const provider = getOAuthProvider(providerId);
      if (!provider) continue;

      const apiProvider = providerMap[providerId] || providerId;

      if (Date.now() >= credentials.expires) {
        log(`  Refreshing expired ${provider.name} token...`, "info");
        try {
          const refreshed = await provider.refreshToken(credentials);
          localStorage.setItem(`oauth_${providerId}`, JSON.stringify(refreshed));
          const apiKey = provider.getApiKey(refreshed);
          await providerKeys.set(apiProvider, apiKey);
          log(`  ✅ ${provider.name} token refreshed`, "ok");
        } catch (e: any) {
          log(`  ⚠️ ${provider.name} refresh failed. Please login again.`, "err");
        }
      } else {
        const apiKey = provider.getApiKey(credentials);
        await providerKeys.set(apiProvider, apiKey);
        const mins = Math.round((credentials.expires - Date.now()) / 60000);
        log(`  ✅ ${provider.name} session restored (${mins}m remaining)`, "ok");
      }
    } catch (e: any) {
      log(`  ⚠️ Failed to restore ${providerId}: ${e.message}`, "err");
    }
  }
}

async function showOAuthMenu() {
  try {
    const { getOAuthProviders } = await import("@mariozechner/pi-ai");
    const allProviders = getOAuthProviders();
    const providers = allProviders.filter(p => BROWSER_OAUTH_PROVIDERS.includes(p.id));

    const buttonsHtml = providers.map(p =>
      `<button class="oauth-btn" data-id="${p.id}" style="display:block;width:100%;padding:12px;margin:6px 0;border:1px solid #ddd;border-radius:8px;background:#fafafa;cursor:pointer;font-size:14px;text-align:left;">${p.name}</button>`
    ).join("");

    const overlay = document.createElement("div");
    overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;";
    const modal = document.createElement("div");
    modal.style.cssText = "background:#fff;border-radius:12px;padding:24px;max-width:400px;width:90%;box-shadow:0 8px 32px rgba(0,0,0,0.2);font-family:system-ui;";
    modal.innerHTML = `
      <h3 style="margin:0 0 8px;font-size:18px;">🔑 Login with your subscription</h3>
      <p style="margin:0 0 16px;color:#555;font-size:13px;">Use your existing Pro/Max subscription — no API costs.</p>
      ${buttonsHtml}
      <button id="modal-cancel" style="display:block;width:100%;padding:10px;margin-top:12px;border:1px solid #ddd;border-radius:8px;background:#fff;cursor:pointer;color:#888;font-size:13px;">Cancel</button>
    `;
    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    const cleanup = () => { document.body.removeChild(overlay); };
    overlay.addEventListener("click", (e) => { if (e.target === overlay) cleanup(); });
    modal.querySelector("#modal-cancel")!.addEventListener("click", cleanup);

    // Handle provider button clicks
    modal.querySelectorAll(".oauth-btn").forEach(btn => {
      btn.addEventListener("click", async () => {
        const id = (btn as HTMLElement).dataset.id!;
        cleanup();
        await (window as any).oauthLogin(id);
      });
    });
  } catch (e: any) {
    log(`OAuth menu error: ${e.message}`, "err");
  }
}

async function doOAuthLogin(
  providerId: string,
  providerKeys: any, // ProviderKeysStore
) {
  const { getOAuthProvider } = await import("@mariozechner/pi-ai");

  const provider = getOAuthProvider(providerId);
  if (!provider) {
    log(`Unknown OAuth provider: ${providerId}`, "err");
    return;
  }

  log(`Starting ${provider.name} login...`, "info");

  try {
    const credentials = await provider.login({
      onAuth: (info) => {
        // Open auth URL in a new tab
        window.open(info.url, "_blank");
        log(`  Opened login page in new tab. Complete the login there.`, "info");
        if (info.instructions) {
          log(`  ${info.instructions}`, "info");
        }
      },
      onPrompt: async (promptInfo) => {
        const input = await showInputModal(
          "Paste Authorization Code",
          promptInfo.message,
          promptInfo.placeholder,
        );
        if (input === null && !promptInfo.allowEmpty) {
          throw new Error("Login cancelled");
        }
        return input || "";
      },
      onProgress: (message) => {
        log(`  ${message}`, "info");
      },
      onManualCodeInput: async () => {
        const input = await showInputModal(
          "Paste Callback URL",
          "After logging in, the page may show an error (that's normal).<br/>Copy the <b>full URL</b> from your browser's address bar and paste it here.",
          "http://localhost:1455/auth/callback?code=...",
        );
        if (!input) throw new Error("Login cancelled");
        return input;
      },
    });

    // Store credentials - map OAuth provider to API provider
    // OAuth tokens need to be stored differently than API keys
    const apiKey = provider.getApiKey(credentials);

    // Map OAuth provider ID to the API provider name
    const providerMap: Record<string, string> = {
      "anthropic": "anthropic",
      "openai-codex": "openai-codex",
      "github-copilot": "github-copilot",
      "gemini-cli": "google",
      "antigravity": "google",
    };

    const apiProvider = providerMap[providerId] || providerId;
    await providerKeys.set(apiProvider, apiKey);

    // Also store the full credentials for token refresh
    const credsKey = `oauth_${providerId}`;
    localStorage.setItem(credsKey, JSON.stringify(credentials));

    // Store expiry info for auto-refresh
    localStorage.setItem(`oauth_expires_${providerId}`, String(credentials.expires));

    log(`✅ Logged in to ${provider.name} successfully!`, "ok");
    log(`  API key stored for "${apiProvider}" provider. You can now chat!`, "info");
  } catch (e: any) {
    if (e.message === "Login cancelled") {
      log(`  Login cancelled`, "info");
    } else {
      log(`  Login failed: ${e.message}`, "err");
    }
  }
}

async function testExcelTool() {
  const id = "btn-tool";
  log(`TODO: End-to-end Excel tool test (read → LLM → write)`, "info");
  log(`  This test requires a working ChatPanel + API key. Use the chat above.`, "info");
  markButton(id, true);
}

// ============================================================================
// EXPOSE TO WINDOW (for onclick handlers in HTML)
// ============================================================================

Object.assign(window, {
  testReadCells,
  testWriteCells,
  testOverview,
  testSelection,
  testPrecedents,
  testEvents,
  testNamedRanges,
  testConditionalFormat,
  testIndexedDB,
  testCORS,
  testWASM,
  testPyodide,
  testChatPanel,
  testExcelTool,
  showOAuthMenu,
});
