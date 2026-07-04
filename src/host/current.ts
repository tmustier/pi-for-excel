/** Current host singleton for app code after boot resolution. */

import { BrowserHost } from "./browser-host.js";
import { detectSpreadsheetHost } from "./detection.js";
import { OfficeHost } from "./office-host.js";
import type { SpreadsheetHost, SpreadsheetHostKind } from "./types.js";
import { WpsHost } from "./wps-host.js";

let currentHost: SpreadsheetHost | null = null;

export function createSpreadsheetHost(kind: SpreadsheetHostKind = detectSpreadsheetHost()): SpreadsheetHost {
  switch (kind) {
    case "office":
      return new OfficeHost();
    case "wps":
      return new WpsHost();
    case "browser":
      return new BrowserHost();
  }
}

export function setCurrentSpreadsheetHost(host: SpreadsheetHost): void {
  currentHost = host;
}

export function getCurrentSpreadsheetHost(): SpreadsheetHost {
  return currentHost ?? createSpreadsheetHost();
}

export function resetCurrentSpreadsheetHostForTests(): void {
  currentHost = null;
}
