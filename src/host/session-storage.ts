/** SettingsStore-backed session/workbook association adapter shared by hosts. */

import type { SpreadsheetHostSessionStorage } from "./types.js";
import {
  getLatestSessionForWorkbook,
  getSessionWorkbookId,
  linkSessionToWorkbook,
  partitionSessionIdsByWorkbook,
  setLatestSessionForWorkbook,
} from "../workbook/session-association.js";

export const settingsBackedSessionStorage: SpreadsheetHostSessionStorage = {
  getSessionWorkbookId,
  linkSessionToWorkbook,
  setLatestSessionForWorkbook,
  getLatestSessionForWorkbook,
  partitionSessionIdsByWorkbook,
};
