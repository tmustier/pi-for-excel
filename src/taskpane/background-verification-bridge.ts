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

import { getAppStorage } from "../storage/local/app-storage.js";
import { closeOverlayById } from "../ui/overlay-dialog.js";
import { MODEL_SELECTOR_OVERLAY_ID } from "../ui/overlay-ids.js";
import type { PiSidebar } from "../ui/pi-sidebar.js";
import type { WorkbookContext } from "../workbook/context.js";
import {
  executeExtensionVerificationCommand,
  isExtensionVerificationCommand,
  type ExtensionVerificationCommandType,
  type ExtensionVerificationOptions,
} from "./background-extension-verification.js";
import type { SessionRuntime } from "./session-runtime-manager.js";

type CoreBridgeCommandType =
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
  | "submitPrompt"
  | "listCharts";

type BridgeCommandType = CoreBridgeCommandType | ExtensionVerificationCommandType;

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

interface BridgeOptions extends ExtensionVerificationOptions {
  sidebar: PiSidebar;
  getWorkbookContext: () => Promise<WorkbookContext>;
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
    isStreaming: runtime.agent.state.isStreaming,
    isBusy: isRuntimeBusy(runtime),
  };
}

async function waitForRuntimeIdle(
  getActiveRuntime: () => SessionRuntime | null,
  initialMessageCount: number,
  timeoutMs: number,
): Promise<JsonRecord> {
  const started = Date.now();
  let sawProgress = false;
  while (Date.now() - started < timeoutMs) {
    const runtime = getActiveRuntime();
    const messageCount = runtime?.agent.state.messages.length ?? 0;
    const busy = isRuntimeBusy(runtime);
    sawProgress ||= messageCount > initialMessageCount || busy;
    if (sawProgress && !busy) {
      return {
        idle: true,
        elapsedMs: Date.now() - started,
        activeRuntime: activeRuntimeSummary(runtime),
      };
    }
    await new Promise((resolve) => window.setTimeout(resolve, 250));
  }

  return {
    idle: false,
    elapsedMs: Date.now() - started,
    activeRuntime: activeRuntimeSummary(getActiveRuntime()),
  };
}

async function waitForBridgeValue<T>(
  readValue: () => T | null,
  description: string,
  timeoutMs = 10_000,
): Promise<T> {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const value = readValue();
    if (value !== null) return value;
    await new Promise((resolve) => window.setTimeout(resolve, 50));
  }
  throw new Error(`Timed out waiting for ${description}`);
}

function selectorRowIdentity(row: HTMLButtonElement): { provider: string; id: string } | null {
  const provider = row.querySelector<HTMLElement>(".pi-model-selector-item-provider")?.textContent?.trim();
  const id = row.querySelector<HTMLElement>(".pi-model-selector-item-id")?.textContent?.trim();
  if (!provider || !id) return null;
  return { provider, id };
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

async function selectModel(payload: DynamicValue, options: BridgeOptions): Promise<JsonRecord> {
  const provider = stringField(payload, "provider");
  const modelId = stringField(payload, "modelId");
  if (!provider || !modelId) {
    throw new Error("selectModel requires payload.provider and payload.modelId");
  }

  const runtime = options.getActiveRuntime();
  if (!runtime) throw new Error("Cannot select a model without an active runtime");
  if (isRuntimeBusy(runtime)) throw new Error("Cannot select a model while the active runtime is busy");

  const before = activeRuntimeSummary(runtime);
  closeOverlayById(MODEL_SELECTOR_OVERLAY_ID);

  const modelButton = await waitForBridgeValue(
    () => document.querySelector<HTMLButtonElement>(".pi-status-model"),
    "the status-bar model button",
  );
  modelButton.click();

  const searchInput = await waitForBridgeValue(
    () => document.querySelector<HTMLInputElement>(".pi-model-selector-search"),
    "the model selector",
  );
  searchInput.value = modelId;
  searchInput.dispatchEvent(new Event("input", { bubbles: true }));

  const row = await waitForBridgeValue(() => {
    const rows = Array.from(
      document.querySelectorAll<HTMLButtonElement>(".pi-model-selector-item"),
    );
    return rows.find((candidate) => {
      const identity = selectorRowIdentity(candidate);
      return identity?.provider === provider && identity.id === modelId;
    }) ?? null;
  }, `${provider}/${modelId} in the model selector`);

  const identity = selectorRowIdentity(row);
  const selectorText = row.textContent?.replace(/\s+/gu, " ").trim() ?? "";
  row.click();

  await waitForBridgeValue(
    () => document.getElementById(MODEL_SELECTOR_OVERLAY_ID) === null ? true : null,
    "the model selector to close",
  );

  const after = await waitForBridgeValue(() => {
    const activeRuntime = options.getActiveRuntime();
    if (
      activeRuntime?.agent.state.model.provider !== provider
      || activeRuntime.agent.state.model.id !== modelId
    ) {
      return null;
    }
    return activeRuntimeSummary(activeRuntime);
  }, `${provider}/${modelId} to become the active model`);

  return {
    selected: true,
    requested: { provider, modelId },
    selectorMatch: {
      provider: identity?.provider ?? "",
      id: identity?.id ?? "",
      text: selectorText,
    },
    before,
    after,
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
  const timeoutMs = Math.max(1_000, Math.min(120_000, numberField(payload, "timeoutMs") ?? 60_000));
  const before = activeRuntimeSummary(runtime);
  const initialMessageCount = runtime?.agent.state.messages.length ?? 0;

  options.sidebar.sendMessage(text);

  return {
    submitted: true,
    textLength: text.length,
    before,
    wait: waitForIdle ? await waitForRuntimeIdle(options.getActiveRuntime, initialMessageCount, timeoutMs) : null,
    after: activeRuntimeSummary(options.getActiveRuntime()),
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
  if (isExtensionVerificationCommand(command.type)) {
    return await executeExtensionVerificationCommand(
      command.type,
      command.payload ?? null,
      options,
    );
  }

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
    case "submitPrompt":
      return await submitPrompt(command.payload, options);
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
