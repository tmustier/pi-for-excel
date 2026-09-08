/**
 * Builtin command overlays (aggregator).
 */

export {
  openSettings,
  configureSettingsPages,
  type SettingsPageId,
  type ExtensionsHubTab,
} from "./settings-pages/index.js";
export { showResumeDialog } from "./resume-overlay.js";
export type {
  RecoveryCheckpointSummary,
  RecoveryCheckpointToolName,
} from "./settings-pages/backups-page.js";
