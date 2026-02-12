/**
 * conventions — public API for the composable cell styles system.
 */

// Types
export type {
  BorderWeight,
  NumberPreset,
  NumberFormatConventions,
  CellStyle,
  NamedStyle,
  ResolvedCellStyle,
  StoredConventions,
  ResolvedConventions,
} from "./types.js";

// Defaults
export {
  DEFAULT_CONVENTIONS,
  DEFAULT_CURRENCY_SYMBOL,
  PRESET_DEFAULT_DP,
  DEFAULT_CONVENTION_CONFIG,
  BUILTIN_STYLES,
  BUILTIN_STYLE_NAMES,
  FORMAT_PRESET_NAMES,
} from "./defaults.js";

// Format builder
export { buildFormatString, isPresetName } from "./format-builder.js";
export type { FormatBuildResult } from "./format-builder.js";

// Style resolver
export { resolveStyles } from "./style-resolver.js";

// Store
export {
  getStoredConventions,
  setStoredConventions,
  resolveConventions,
  getResolvedConventions,
  mergeStoredConventions,
  diffFromDefaults,
} from "./store.js";
export type { ConventionsStore, ConventionDiff } from "./store.js";

// Humanize
export { humanizeFormat } from "./humanize.js";
