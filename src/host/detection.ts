import type { SpreadsheetHostKind } from "./types.js";

interface HostGlobalScope {
  readonly wps?: unknown;
  readonly Application?: unknown;
  readonly Office?: unknown;
}

function getGlobalMember(scope: object, key: keyof HostGlobalScope): unknown {
  const hostGlobals = scope as HostGlobalScope;
  return hostGlobals[key];
}

function hasObjectOrFunction(value: unknown): value is object {
  return (typeof value === "object" && value !== null) || typeof value === "function";
}

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

function readMember(owner: object, key: string): unknown {
  try {
    return Reflect.get(owner, key);
  } catch {
    return undefined;
  }
}

function looksLikeWpsEtApplication(value: unknown): boolean {
  if (!hasObjectOrFunction(value)) return false;
  return WPS_ET_APPLICATION_MEMBERS.some((key) => readMember(value, key) !== undefined);
}

function looksLikeWpsPluginGlobal(value: unknown): boolean {
  if (!hasObjectOrFunction(value)) return false;
  if (typeof readMember(value, "EtApplication") === "function") return true;
  if (readMember(value, "PluginStorage") !== undefined) return true;
  return looksLikeWpsEtApplication(readMember(value, "Application"));
}

export function hasWpsJsApiGlobal(scope: object = globalThis): boolean {
  if (looksLikeWpsPluginGlobal(getGlobalMember(scope, "wps"))) return true;
  return looksLikeWpsEtApplication(getGlobalMember(scope, "Application"));
}

export function hasOfficeJsGlobal(scope: object = globalThis): boolean {
  return hasObjectOrFunction(getGlobalMember(scope, "Office"));
}

export function detectSpreadsheetHost(scope: object = globalThis): SpreadsheetHostKind {
  if (hasWpsJsApiGlobal(scope)) return "wps";
  if (hasOfficeJsGlobal(scope)) return "office";
  return "browser";
}
