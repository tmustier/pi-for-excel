export type {
  SpreadsheetHost,
  SpreadsheetHostKind,
  SpreadsheetHostReadyCallback,
  SpreadsheetHostReadyInfo,
  SpreadsheetHostSessionStorage,
} from "./types.js";
export type { WorkbookContext } from "./workbook-context.js";
export { formatWorkbookLabel } from "./workbook-context.js";
export {
  detectSpreadsheetHost,
  hasOfficeJsGlobal,
  hasWpsJsApiGlobal,
} from "./detection.js";
export {
  createSpreadsheetHost,
  getCurrentSpreadsheetHost,
  resetCurrentSpreadsheetHostForTests,
  setCurrentSpreadsheetHost,
} from "./current.js";
export { resolveSpreadsheetHostForBoot } from "./boot.js";
export { BrowserHost } from "./browser-host.js";
export { OfficeHost } from "./office-host.js";
export { WpsHost, WPS_UNSUPPORTED_PHASE_1_MESSAGE } from "./wps-host.js";
