/**
 * Dev-only background verification bridge.
 *
 * Enabled only when the Vite dev build provides both:
 * - VITE_PI_BACKGROUND_VERIFY_URL (for example https://localhost:3157)
 * - VITE_PI_BACKGROUND_VERIFY_TOKEN
 *
 * The taskpane initiates all network traffic to a tokened loopback server.
 * This lets agents verify the real Excel host and real taskpane while Excel
 * remains in the background; no raw GUI input is required.
 */

import type { ThinkingLevel } from "@earendil-works/pi-agent-core";
import type { Api, Model } from "@earendil-works/pi-ai/compat";

import type { WorkbookContext } from "../workbook/context.js";
import { getAppStorage } from "../storage/local/app-storage.js";
import type { PiSidebar } from "../ui/pi-sidebar.js";
import { decideRuntimeIdle } from "./background-verify-idle.js";
import {
  collectBuiltInModelCandidates,
  resolveBridgeModelSelection,
  type BridgeModelCandidate,
} from "./background-verify-model.js";
import {
  assistantTextSnippet,
  buildTranscriptExport,
  summarizeLastToolCall,
  type TranscriptExportOptions,
} from "./background-verify-transcript.js";
import type { SessionRuntime } from "./session-runtime-manager.js";

type BridgeCommandType =
  | "noop"
  | "status"
  | "officeProbe"
  | "readRange"
  | "readUsedRange"
  | "writeRange"
  | "clearRange"
  | "workbookWriteProbe"
  | "configureProxy"
  | "selectModel"
  | "newSession"
  | "submitPrompt"
  | "waitUntilIdle"
  | "exportTranscript"
  | "listCharts";

interface BridgeCommand {
  id?: string;
  type: BridgeCommandType;
  payload?: DynamicValue;
}

interface PollResponse extends BridgeCommand {
  error?: string;
}

interface BridgeClientRegistration {
  clientId: string;
}

interface BridgeOptions {
  sidebar: PiSidebar;
  getWorkbookContext: () => Promise<WorkbookContext>;
  getActiveRuntime: () => SessionRuntime | null;
  /** Production fresh-session seam (new chat tab/runtime). */
  createNewSession: () => Promise<SessionRuntime | null>;
  /** Production model-switch seam: applies a registry model + thinking level in place. */
  selectRuntimeModel: (args: {
    runtimeId: string;
    model: Model<Api>;
    thinkingLevel: ThinkingLevel;
  }) => Promise<void>;
}

interface JsonRecord {
  [key: string]: DynamicValue;
}

interface BridgeStopHandle {
  stop: () => void;
}

const DEFAULT_POLL_DELAY_MS = 750;

function envValue(name: keyof ImportMetaEnv): string {
  const value: DynamicValue = import.meta.env[name];
  return typeof value === "string" ? value.trim() : "";
}

function isTaskpaneBackgroundVerificationBridgePayloadShape(value: DynamicValue): value is JsonRecord {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function stringField(value: DynamicValue, key: string): string | undefined {
  if (!isTaskpaneBackgroundVerificationBridgePayloadShape(value)) return undefined;
  const field = value[key];
  return typeof field === "string" && field.trim().length > 0 ? field.trim() : undefined;
}

function booleanField(value: DynamicValue, key: string): boolean | undefined {
  if (!isTaskpaneBackgroundVerificationBridgePayloadShape(value)) return undefined;
  const field = value[key];
  return typeof field === "boolean" ? field : undefined;
}

function numberField(value: DynamicValue, key: string): number | undefined {
  if (!isTaskpaneBackgroundVerificationBridgePayloadShape(value)) return undefined;
  const field = value[key];
  return typeof field === "number" && Number.isFinite(field) ? field : undefined;
}

function isUnknownArray(value: DynamicValue): value is readonly DynamicValue[] {
  return Array.isArray(value);
}

function matrixField(value: DynamicValue, key: string): DynamicValue[][] | undefined {
  if (!isTaskpaneBackgroundVerificationBridgePayloadShape(value)) return undefined;
  const field = value[key];
  if (!isUnknownArray(field) || field.length === 0) return undefined;
  const rows: DynamicValue[][] = [];
  let width: number | null = null;
  for (const row of field) {
    if (!isUnknownArray(row) || row.length === 0) return undefined;
    width ??= row.length;
    if (row.length !== width) return undefined;
    rows.push(row.map((cell) => cell));
  }
  return rows;
}

function bridgeUrl(): string | null {
  const configured = envValue("VITE_PI_BACKGROUND_VERIFY_URL");
  if (!configured) return null;

  try {
    const parsed = new URL(configured);
    if (parsed.protocol !== "https:") return null;
    if (parsed.hostname !== "localhost" && parsed.hostname !== "127.0.0.1") return null;
    return parsed.toString().replace(/\/$/u, "");
  } catch {
    return null;
  }
}

function bridgeToken(): string | null {
  const configured = envValue("VITE_PI_BACKGROUND_VERIFY_TOKEN");
  return configured.length > 0 ? configured : null;
}

async function postJson<T>(url: string, body: JsonRecord, signal: AbortSignal): Promise<T> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify(body),
    signal,
  });
  const parsed = await response.json() as DynamicValue;
  if (!response.ok) {
    const message = stringField(parsed, "error") ?? `HTTP ${response.status}`;
    throw new Error(message);
  }
  return parsed as T;
}

function assertNever(value: never): never {
  throw new Error(`Unknown background verification command: ${String(value)}`);
}

function serializeError(error: DynamicValue): JsonRecord {
  if (error instanceof Error) {
    return {
      name: error.name,
      message: error.message,
    };
  }
  return { message: String(error) };
}

function isRuntimeBusy(runtime: SessionRuntime | null): boolean {
  return runtime ? runtime.agent.state.isStreaming || runtime.actionQueue.isBusy() : false;
}

function latestAssistantSummary(runtime: SessionRuntime): JsonRecord | null {
  for (let index = runtime.agent.state.messages.length - 1; index >= 0; index -= 1) {
    const message = runtime.agent.state.messages[index];
    if (!message || message.role !== "assistant") continue;

    const textLength = message.content.reduce((total, part) => (
      part.type === "text" ? total + part.text.length : total
    ), 0);
    return {
      provider: message.provider,
      model: message.model,
      api: message.api,
      stopReason: message.stopReason,
      errorMessage: message.errorMessage,
      textLength,
      snippet: assistantTextSnippet(message),
      usage: message.usage,
    };
  }
  return null;
}

function activeRuntimeSummary(runtime: SessionRuntime | null): JsonRecord | null {
  if (!runtime) return null;
  return {
    runtimeId: runtime.runtimeId,
    sessionId: runtime.agent.sessionId,
    model: runtime.agent.state.model,
    thinkingLevel: runtime.agent.state.thinkingLevel,
    messageCount: runtime.agent.state.messages.length,
    lastAssistant: latestAssistantSummary(runtime),
    lastToolCall: summarizeLastToolCall(runtime.agent.state.messages),
    isStreaming: runtime.agent.state.isStreaming,
    isBusy: isRuntimeBusy(runtime),
  };
}

const IDLE_POLL_INTERVAL_MS = 250;
// Keep the inline wait under the server command timeout ceiling (300s) so the
// server can still return a bounded 504 rather than hanging.
const MAX_IDLE_WAIT_MS = 285_000;
const DEFAULT_IDLE_WAIT_MS = 60_000;
const DEFAULT_STARTUP_GRACE_MS = 30_000;
const MAX_STARTUP_GRACE_MS = 120_000;

function clampIdleWaitMs(value: number | undefined): number {
  const base = typeof value === "number" && Number.isFinite(value) ? Math.floor(value) : DEFAULT_IDLE_WAIT_MS;
  return Math.max(1_000, Math.min(MAX_IDLE_WAIT_MS, base));
}

function clampStartupGraceMs(value: number | undefined): number {
  const base = typeof value === "number" && Number.isFinite(value) ? Math.floor(value) : DEFAULT_STARTUP_GRACE_MS;
  return Math.max(0, Math.min(MAX_STARTUP_GRACE_MS, base));
}

/**
 * Wait until the active runtime has started (busy, or message count grew past
 * the pre-run baseline) and then returned to idle. Bounded by an explicit
 * timeout and a startup grace so a run that never starts (dead proxy/OAuth)
 * surfaces as `start-timeout` instead of a false immediate idle.
 */
async function waitForRuntimeIdle(
  getActiveRuntime: () => SessionRuntime | null,
  baselineMessageCount: number,
  timeoutMs: number,
  startupGraceMs: number,
): Promise<JsonRecord> {
  const started = Date.now();
  let sawStart = false;

  while (Date.now() - started < timeoutMs) {
    const runtime = getActiveRuntime();
    const decision = decideRuntimeIdle({
      baselineMessageCount,
      currentMessageCount: runtime?.agent.state.messages.length ?? 0,
      isBusy: isRuntimeBusy(runtime),
      sawStart,
      observedElapsedMs: Date.now() - started,
      startupGraceMs,
    });
    sawStart = sawStart || decision.started;
    if (decision.done) {
      return {
        idle: decision.idle,
        started: decision.started,
        reason: decision.reason,
        elapsedMs: Date.now() - started,
        baselineMessageCount,
        activeRuntime: activeRuntimeSummary(runtime),
      };
    }
    await new Promise((resolve) => window.setTimeout(resolve, IDLE_POLL_INTERVAL_MS));
  }

  return {
    idle: false,
    started: sawStart,
    reason: "wait-timeout",
    elapsedMs: Date.now() - started,
    baselineMessageCount,
    activeRuntime: activeRuntimeSummary(getActiveRuntime()),
  };
}

async function configureProxy(payload: DynamicValue): Promise<JsonRecord> {
  const enabled = booleanField(payload, "enabled");
  if (enabled === undefined) {
    throw new Error("configureProxy requires boolean payload.enabled");
  }

  const url = stringField(payload, "url");
  const settings = getAppStorage().settings;
  await settings.set("proxy.enabled", enabled);
  if (url) {
    await settings.set("proxy.url", url);
  }

  return {
    configured: true,
    enabled,
    url: url ?? null,
  };
}

async function collectModelCandidates(): Promise<BridgeModelCandidate[]> {
  const candidates = collectBuiltInModelCandidates();
  try {
    const customProviders = await getAppStorage().customProviders.getAll();
    for (const provider of customProviders) {
      if (!provider.models) continue;
      for (const model of provider.models) {
        candidates.push({ provider: model.provider, id: model.id, model });
      }
    }
  } catch (error) {
    console.warn("[pi] Background verify: failed to load custom provider models", error);
  }
  return candidates;
}

async function selectModel(payload: DynamicValue, options: BridgeOptions): Promise<JsonRecord> {
  const provider = stringField(payload, "provider");
  const modelId = stringField(payload, "modelId");
  if (!provider || !modelId) {
    throw new Error("selectModel requires payload.provider and payload.modelId");
  }
  const requestedThinkingLevel = stringField(payload, "thinkingLevel");

  const runtime = options.getActiveRuntime();
  if (!runtime) throw new Error("Cannot select a model without an active runtime");
  if (isRuntimeBusy(runtime)) throw new Error("Cannot select a model while the active runtime is busy");

  const before = activeRuntimeSummary(runtime);
  const candidates = await collectModelCandidates();
  const resolution = resolveBridgeModelSelection({
    candidates,
    provider,
    modelId,
    requestedThinkingLevel,
  });
  if (!resolution.ok) {
    throw new Error(resolution.error);
  }

  // Apply via the production model-switch seam (registry model + validated
  // thinking level), not a direct agent-state mutation or DOM click hack.
  await options.selectRuntimeModel({
    runtimeId: runtime.runtimeId,
    model: resolution.model,
    thinkingLevel: resolution.thinkingLevel,
  });

  const applied = options.getActiveRuntime();
  const appliedModel = applied?.agent.state.model;
  if (!appliedModel || appliedModel.provider !== resolution.provider || appliedModel.id !== resolution.modelId) {
    throw new Error(`Model switch did not apply ${resolution.provider}/${resolution.modelId}`);
  }
  if (applied?.agent.state.thinkingLevel !== resolution.thinkingLevel) {
    throw new Error(`Thinking level did not apply ${resolution.thinkingLevel}`);
  }

  return {
    selected: true,
    requested: { provider, modelId, thinkingLevel: requestedThinkingLevel ?? null },
    resolved: {
      provider: resolution.provider,
      modelId: resolution.modelId,
      thinkingLevel: resolution.thinkingLevel,
      requestedThinkingLevel: resolution.requestedThinkingLevel,
    },
    supportedThinkingLevels: resolution.supportedThinkingLevels,
    before,
    after: activeRuntimeSummary(applied),
  };
}

async function submitPrompt(payload: DynamicValue, options: BridgeOptions): Promise<JsonRecord> {
  const text = stringField(payload, "text");
  if (!text) throw new Error("submitPrompt requires payload.text");

  const runtime = options.getActiveRuntime();
  if (isRuntimeBusy(runtime)) {
    throw new Error("Cannot submit prompt while the active runtime is busy");
  }

  const waitForIdle = booleanField(payload, "waitForIdle") ?? true;
  const timeoutMs = clampIdleWaitMs(numberField(payload, "timeoutMs"));
  const startupGraceMs = clampStartupGraceMs(numberField(payload, "startupGraceMs"));
  const before = activeRuntimeSummary(runtime);
  const baselineMessageCount = runtime?.agent.state.messages.length ?? 0;

  options.sidebar.sendMessage(text);

  return {
    submitted: true,
    textLength: text.length,
    baseline: {
      messageCount: baselineMessageCount,
      runtimeId: runtime?.runtimeId ?? null,
      sessionId: runtime?.agent.sessionId ?? null,
    },
    before,
    wait: waitForIdle
      ? await waitForRuntimeIdle(options.getActiveRuntime, baselineMessageCount, timeoutMs, startupGraceMs)
      : null,
    after: activeRuntimeSummary(options.getActiveRuntime()),
  };
}

/**
 * Pollable wait-until-idle. Pass `baselineMessageCount` from a prior
 * `submitPrompt` response for exact start detection; when omitted the current
 * message count is used and only a live busy state counts as "started", which
 * fails closed rather than reporting a false idle.
 */
async function waitUntilIdle(payload: DynamicValue, options: BridgeOptions): Promise<JsonRecord> {
  const runtime = options.getActiveRuntime();
  if (!runtime) throw new Error("waitUntilIdle requires an active runtime");

  const timeoutMs = clampIdleWaitMs(numberField(payload, "timeoutMs"));
  const startupGraceMs = clampStartupGraceMs(numberField(payload, "startupGraceMs"));
  const explicitBaseline = numberField(payload, "baselineMessageCount");
  const baselineMessageCount = typeof explicitBaseline === "number" && Number.isFinite(explicitBaseline)
    ? Math.max(0, Math.floor(explicitBaseline))
    : runtime.agent.state.messages.length;

  return waitForRuntimeIdle(options.getActiveRuntime, baselineMessageCount, timeoutMs, startupGraceMs);
}

async function startNewSession(options: BridgeOptions): Promise<JsonRecord> {
  const previous = options.getActiveRuntime();
  if (isRuntimeBusy(previous)) {
    throw new Error("Cannot start a new session while the active runtime is busy");
  }

  const before = activeRuntimeSummary(previous);
  const created = await options.createNewSession();
  if (!created) throw new Error("Failed to create a new background-verification session");

  const after = activeRuntimeSummary(options.getActiveRuntime());
  const beforeRuntimeId = previous?.runtimeId ?? null;
  const afterRuntimeId = options.getActiveRuntime()?.runtimeId ?? null;
  const beforeSessionId = previous?.agent.sessionId ?? null;
  const afterSessionId = options.getActiveRuntime()?.agent.sessionId ?? null;

  return {
    created: true,
    before,
    after,
    activeRuntimeId: afterRuntimeId,
    runtimeChanged: beforeRuntimeId !== afterRuntimeId,
    sessionChanged: beforeSessionId !== afterSessionId,
    messageCountBefore: previous?.agent.state.messages.length ?? null,
    messageCountAfter: options.getActiveRuntime()?.agent.state.messages.length ?? null,
  };
}

function transcriptOptionsFromPayload(payload: DynamicValue): TranscriptExportOptions {
  const options: TranscriptExportOptions = {};
  const maxReplyChars = numberField(payload, "maxReplyChars");
  if (maxReplyChars !== undefined) options.maxReplyChars = maxReplyChars;
  const maxMessages = numberField(payload, "maxMessages");
  if (maxMessages !== undefined) options.maxMessages = maxMessages;
  const maxMessageTextChars = numberField(payload, "maxMessageTextChars");
  if (maxMessageTextChars !== undefined) options.maxMessageTextChars = maxMessageTextChars;
  const maxTools = numberField(payload, "maxTools");
  if (maxTools !== undefined) options.maxTools = maxTools;
  return options;
}

function exportTranscript(payload: DynamicValue, options: BridgeOptions): JsonRecord {
  const runtime = options.getActiveRuntime();
  if (!runtime) throw new Error("exportTranscript requires an active runtime");

  const transcript = buildTranscriptExport(
    runtime.agent.state.messages,
    transcriptOptionsFromPayload(payload),
  );

  return {
    runtimeId: runtime.runtimeId,
    sessionId: runtime.agent.sessionId,
    model: runtime.agent.state.model,
    thinkingLevel: runtime.agent.state.thinkingLevel,
    isBusy: isRuntimeBusy(runtime),
    transcript,
  };
}

async function runOfficeProbe(): Promise<JsonRecord> {
  if (typeof Excel === "undefined") {
    throw new Error("Excel global is unavailable; the taskpane is not running inside the Excel host.");
  }

  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    const worksheets = workbook.worksheets;
    const activeWorksheet = worksheets.getActiveWorksheet();
    const selectedRange = workbook.getSelectedRange();

    worksheets.load("items/name");
    activeWorksheet.load("name");
    selectedRange.load("address,rowCount,columnCount,values,text,formulas");
    await context.sync();

    return {
      activeWorksheet: activeWorksheet.name,
      worksheetNames: worksheets.items.map((sheet) => sheet.name),
      selectedRange: {
        address: selectedRange.address,
        rowCount: selectedRange.rowCount,
        columnCount: selectedRange.columnCount,
        values: selectedRange.values,
        text: selectedRange.text,
        formulas: selectedRange.formulas,
      },
    };
  });
}

function parseQualifiedRange(address: string): { sheetName?: string; rangeAddress: string } {
  const trimmed = address.trim();
  const bang = trimmed.lastIndexOf("!");
  if (bang < 0) return { rangeAddress: trimmed };

  let sheetName = trimmed.slice(0, bang).trim();
  if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
    sheetName = sheetName.slice(1, -1).replaceAll("''", "'");
  }
  return { sheetName, rangeAddress: trimmed.slice(bang + 1).trim() };
}

function summarizeRange(range: Excel.Range): JsonRecord {
  return {
    address: range.address,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    values: range.values,
    text: range.text,
    formulas: range.formulas,
    numberFormat: range.numberFormat,
  };
}

async function readRange(address: string): Promise<JsonRecord> {
  if (typeof Excel === "undefined") {
    throw new Error("Excel global is unavailable; the taskpane is not running inside the Excel host.");
  }

  return await Excel.run(async (context) => {
    const parsed = parseQualifiedRange(address);
    const sheet = parsed.sheetName
      ? context.workbook.worksheets.getItem(parsed.sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(parsed.rangeAddress);
    range.load("address,rowCount,columnCount,values,text,formulas,numberFormat");
    await context.sync();
    return summarizeRange(range);
  });
}

async function writeRange(address: string, values: DynamicValue[][], formulas?: DynamicValue[][], numberFormat?: DynamicValue[][]): Promise<JsonRecord> {
  if (typeof Excel === "undefined") {
    throw new Error("Excel global is unavailable; the taskpane is not running inside the Excel host.");
  }

  return await Excel.run(async (context) => {
    const parsed = parseQualifiedRange(address);
    const sheet = parsed.sheetName
      ? context.workbook.worksheets.getItem(parsed.sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(parsed.rangeAddress);
    range.load("address,rowCount,columnCount,values,text,formulas,numberFormat");
    await context.sync();
    const before = summarizeRange(range);

    if (formulas) {
      range.formulas = formulas;
    } else {
      range.values = values;
    }
    if (numberFormat) range.numberFormat = numberFormat;
    range.load("address,rowCount,columnCount,values,text,formulas,numberFormat");
    await context.sync();

    return {
      address: range.address,
      before,
      after: summarizeRange(range),
    };
  });
}

function clearApplyToFromPayload(payload: DynamicValue): Excel.ClearApplyTo {
  const applyTo = stringField(payload, "applyTo") ?? "contents";
  switch (applyTo) {
    case "all":
      return Excel.ClearApplyTo.all;
    case "formats":
      return Excel.ClearApplyTo.formats;
    case "contents":
      return Excel.ClearApplyTo.contents;
    default:
      throw new Error("clearRange payload.applyTo must be one of: contents, formats, all");
  }
}

async function clearRange(address: string, applyTo: Excel.ClearApplyTo): Promise<JsonRecord> {
  if (typeof Excel === "undefined") {
    throw new Error("Excel global is unavailable; the taskpane is not running inside the Excel host.");
  }

  return await Excel.run(async (context) => {
    const parsed = parseQualifiedRange(address);
    const sheet = parsed.sheetName
      ? context.workbook.worksheets.getItem(parsed.sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(parsed.rangeAddress);
    range.load("address,rowCount,columnCount,values,text,formulas,numberFormat");
    await context.sync();
    const before = summarizeRange(range);

    range.clear(applyTo);
    range.load("address,rowCount,columnCount,values,text,formulas,numberFormat");
    await context.sync();

    return {
      address: range.address,
      applyTo,
      before,
      after: summarizeRange(range),
    };
  });
}

async function workbookWriteProbe(payload: DynamicValue): Promise<JsonRecord> {
  if (typeof Excel === "undefined") {
    throw new Error("Excel global is unavailable; the taskpane is not running inside the Excel host.");
  }

  const sheetName = stringField(payload, "sheetName") ?? "_pi_background_verify";
  const marker = stringField(payload, "marker") ?? `pi-background-verify-${new Date().toISOString()}`;
  const keepSheet = booleanField(payload, "keepSheet") ?? false;

  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getItemOrNullObject(sheetName);
    sheet.load("name,isNullObject");
    await context.sync();

    const createdSheet = sheet.isNullObject;
    if (createdSheet) {
      sheet = sheets.add(sheetName);
      sheet.load("name");
      await context.sync();
    }

    const target = sheet.getRange("A1:B4");
    target.load("address,rowCount,columnCount,values,text,formulas,numberFormat");
    await context.sync();
    const before = summarizeRange(target);

    target.values = [
      ["marker", marker],
      ["input", 2],
      ["input", 3],
      ["sum", ""],
    ];
    sheet.getRange("B4").formulas = [["=SUM(B2:B3)"]];
    target.load("address,rowCount,columnCount,values,text,formulas,numberFormat");
    await context.sync();
    const afterWrite = summarizeRange(target);

    let cleanup: JsonRecord;
    if (createdSheet && !keepSheet) {
      sheet.delete();
      await context.sync();
      cleanup = { action: "delete-created-sheet", restored: true };
    } else if (!createdSheet && !keepSheet) {
      target.formulas = before.formulas as DynamicValue[][];
      target.numberFormat = before.numberFormat as DynamicValue[][];
      await context.sync();
      cleanup = { action: "restore-existing-range", restored: true };
    } else {
      cleanup = { action: "leave-written-range", restored: false };
    }

    return {
      sheetName,
      marker,
      createdSheet,
      range: target.address,
      before,
      afterWrite,
      cleanup,
    };
  });
}

async function readUsedRange(): Promise<JsonRecord> {
  if (typeof Excel === "undefined") {
    throw new Error("Excel global is unavailable; the taskpane is not running inside the Excel host.");
  }

  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject();
    sheet.load("name");
    used.load("address,rowCount,columnCount,values,text,formulas,isNullObject");
    await context.sync();
    if (used.isNullObject) {
      return { sheetName: sheet.name, usedRange: null };
    }
    return {
      sheetName: sheet.name,
      usedRange: {
        address: used.address,
        rowCount: used.rowCount,
        columnCount: used.columnCount,
        values: used.values,
        text: used.text,
        formulas: used.formulas,
      },
    };
  });
}

async function listCharts(): Promise<JsonRecord> {
  if (typeof Excel === "undefined") {
    throw new Error("Excel global is unavailable; the taskpane is not running inside the Excel host.");
  }

  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    for (const sheet of sheets.items) {
      sheet.charts.load("items/id,items/name,items/chartType,items/top,items/left,items/width,items/height");
    }
    await context.sync();

    return {
      charts: sheets.items.flatMap((sheet) => sheet.charts.items.map((chart) => ({
        sheetName: sheet.name,
        id: chart.id,
        name: chart.name,
        chartType: chart.chartType,
        position: {
          top: chart.top,
          left: chart.left,
          width: chart.width,
          height: chart.height,
        },
      }))),
    };
  });
}

async function executeCommand(command: BridgeCommand, options: BridgeOptions): Promise<DynamicValue> {
  switch (command.type) {
    case "noop":
      return { ok: true };
    case "status": {
      const textarea = options.sidebar.getTextarea();
      return {
        ready: true,
        href: window.location.href,
        origin: window.location.origin,
        visibilityState: document.visibilityState,
        hasOffice: typeof Office !== "undefined",
        hasExcel: typeof Excel !== "undefined",
        workbookContext: await options.getWorkbookContext(),
        activeRuntime: activeRuntimeSummary(options.getActiveRuntime()),
        input: textarea
          ? {
              valueLength: textarea.value.length,
              placeholder: textarea.getAttribute("placeholder") ?? "",
            }
          : null,
      };
    }
    case "officeProbe":
      return await runOfficeProbe();
    case "readRange": {
      const address = stringField(command.payload, "address");
      if (!address) throw new Error("readRange requires payload.address");
      return await readRange(address);
    }
    case "readUsedRange":
      return await readUsedRange();
    case "writeRange": {
      const address = stringField(command.payload, "address");
      const values = matrixField(command.payload, "values");
      const formulas = matrixField(command.payload, "formulas");
      const numberFormat = matrixField(command.payload, "numberFormat");
      if (!address) throw new Error("writeRange requires payload.address");
      if (!values && !formulas) throw new Error("writeRange requires payload.values or payload.formulas");
      return await writeRange(address, values ?? formulas ?? [], formulas, numberFormat);
    }
    case "clearRange": {
      const address = stringField(command.payload, "address");
      if (!address) throw new Error("clearRange requires payload.address");
      return await clearRange(address, clearApplyToFromPayload(command.payload));
    }
    case "workbookWriteProbe":
      return await workbookWriteProbe(command.payload);
    case "configureProxy":
      return await configureProxy(command.payload);
    case "selectModel":
      return await selectModel(command.payload, options);
    case "newSession":
      return await startNewSession(options);
    case "submitPrompt":
      return await submitPrompt(command.payload, options);
    case "waitUntilIdle":
      return await waitUntilIdle(command.payload, options);
    case "exportTranscript":
      return exportTranscript(command.payload, options);
    case "listCharts":
      return await listCharts();
    default:
      return assertNever(command.type);
  }
}

export function maybeStartBackgroundVerificationBridge(options: BridgeOptions): BridgeStopHandle | null {
  if (!import.meta.env.DEV) return null;

  const url = bridgeUrl();
  const token = bridgeToken();
  if (!url || !token) return null;

  const controller = new AbortController();
  const signal = controller.signal;
  const client = {
    href: window.location.href,
    userAgent: navigator.userAgent,
    startedAt: new Date().toISOString(),
  };

  const loop = async (): Promise<void> => {
    let clientId = "";
    while (!signal.aborted) {
      try {
        if (!clientId) {
          const registration = await postJson<BridgeClientRegistration>(
            `${url}/client/register`,
            { token, client },
            signal,
          );
          clientId = registration.clientId;
          console.info("[pi] Background verification bridge connected", { url, clientId });
        }

        const command = await postJson<PollResponse>(
          `${url}/client/poll`,
          { token, clientId },
          signal,
        );
        if (command.type === "noop") continue;

        if (!command.id) {
          console.warn("[pi] Background verification command missing id", command);
          continue;
        }

        try {
          const result = await executeCommand(command, options);
          await postJson(
            `${url}/client/result`,
            { token, clientId, commandId: command.id, ok: true, result },
            signal,
          );
        } catch (error) {
          await postJson(
            `${url}/client/result`,
            { token, clientId, commandId: command.id, ok: false, error: serializeError(error) },
            signal,
          );
        }
      } catch (error) {
        if (signal.aborted) return;
        console.warn("[pi] Background verification bridge disconnected", error);
        clientId = "";
        await new Promise((resolve) => window.setTimeout(resolve, DEFAULT_POLL_DELAY_MS));
      }
    }
  };

  void loop();

  return {
    stop: () => controller.abort(),
  };
}
