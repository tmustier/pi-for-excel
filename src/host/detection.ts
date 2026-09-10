/** Host global detection for taskpane boot. */

import type { SpreadsheetHostKind } from "./types.js";

interface HostGlobalScope {
  readonly wps?: DynamicValue;
  readonly Application?: DynamicValue;
  readonly Office?: DynamicValue;
}

function getGlobalMember(scope: object, key: keyof HostGlobalScope): DynamicValue {
  // Every object is safe to read through this optional, read-only host-probe shape.
  const hostGlobals = scope as HostGlobalScope;
  return hostGlobals[key];
}

function hasObjectOrFunction(value: DynamicValue): value is object {
  return (typeof value === "object" && value !== null) || typeof value === "function";
}

/** Members `src/host/wps/jsapi.ts` reads off the ET `Application` object. */
const WPS_ET_APPLICATION_MEMBERS = [
  "ActiveWorkbook",
  "ActiveSheet",
  "Worksheets",
  "Sheets",
  "Range",
  "PluginStorage",
  "CreateTaskpane",
  "CreateTaskPane",
] as const;

function readMember(owner: object, key: string): DynamicValue {
  try {
    return Reflect.get(owner, key);
  } catch {
    // Host objects may throw from accessors before the document is ready.
    return undefined;
  }
}

function looksLikeWpsEtApplication(value: DynamicValue): boolean {
  if (!hasObjectOrFunction(value)) return false;
  return WPS_ET_APPLICATION_MEMBERS.some((key) => readMember(value, key) !== undefined);
}

function looksLikeWpsPluginGlobal(value: DynamicValue): boolean {
  if (!hasObjectOrFunction(value)) return false;
  if (typeof readMember(value, "EtApplication") === "function") return true;
  if (readMember(value, "PluginStorage") !== undefined) return true;
  return looksLikeWpsEtApplication(readMember(value, "Application"));
}

/**
 * True only when a global carries the WPS JSAPI surface this add-in uses,
 * so an unrelated `Application` or `wps` global does not select the WPS host.
 */
export function hasWpsJsApiGlobal(scope: object = globalThis): boolean {
  if (looksLikeWpsPluginGlobal(getGlobalMember(scope, "wps"))) return true;
  return looksLikeWpsEtApplication(getGlobalMember(scope, "Application"));
}

export function hasOfficeJsGlobal(scope: object = globalThis): boolean {
  return hasObjectOrFunction(getGlobalMember(scope, "Office"));
}

/**
 * Synchronous host detection from globals. Boot resolution may still downgrade
 * `office` to `browser` if Office.onReady never fires within the UI-test timeout.
 */
export function detectSpreadsheetHost(scope: object = globalThis): SpreadsheetHostKind {
  // WPS does not run Office.js; prefer WPS if both globals are ever present.
  if (hasWpsJsApiGlobal(scope)) return "wps";
  if (hasOfficeJsGlobal(scope)) return "office";
  return "browser";
}
