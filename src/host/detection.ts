/** Host global detection for taskpane boot. */

import type { SpreadsheetHostKind } from "./types.js";

function getGlobalMember(scope: object, key: string): DynamicValue {
  return Reflect.get(scope, key);
}

function hasObjectOrFunction(value: DynamicValue): boolean {
  return (typeof value === "object" && value !== null) || typeof value === "function";
}

export function hasWpsJsApiGlobal(scope: object = globalThis): boolean {
  const wps = getGlobalMember(scope, "wps");
  if (hasObjectOrFunction(wps)) return true;

  const application = getGlobalMember(scope, "Application");
  return hasObjectOrFunction(application);
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
