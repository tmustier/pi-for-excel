/**
 * Builtin command for managing experimental feature flags.
 */

import type { SlashCommand } from "../types.js";
import {
  getExperimentalFeatureSlugs,
  isExperimentalFeatureEnabled,
  resolveExperimentalFeature,
  setExperimentalFeatureEnabled,
  toggleExperimentalFeature,
  type ExperimentalFeatureDefinition,
  type ExperimentalFeatureId,
} from "../../experiments/flags.js";
import { validateOfficeProxyUrl } from "../../auth/proxy-validation.js";
import { dispatchExperimentalToolConfigChanged } from "../../experiments/events.js";
import {
  buildTmuxBridgeGateErrorMessage,
  PYTHON_BRIDGE_URL_SETTING_KEY,
  TMUX_BRIDGE_URL_SETTING_KEY,
  type TmuxBridgeGateReason,
  type TmuxBridgeGateResult,
} from "../../tools/experimental-tool-gates.js";
import { PYTHON_BRIDGE_TOKEN_SETTING_KEY } from "../../tools/python-run.js";
import { TMUX_BRIDGE_TOKEN_SETTING_KEY } from "../../tools/tmux.js";
import { isRecord } from "../../utils/type-guards.js";
import { showToast } from "../../ui/toast.js";
import { showExperimentalDialog } from "./experimental-overlay.js";

const ENABLE_ACTIONS = new Set(["enable", "on"]);
const DISABLE_ACTIONS = new Set(["disable", "off"]);
const TOGGLE_ACTIONS = new Set(["toggle"]);
const OPEN_ACTIONS = new Set(["open", "ui", "list", "status"]);

const LEGACY_INTEGRATIONS_FEATURES = new Set(["mcp", "mcp-tools", "external-tools"]);

const TMUX_BRIDGE_URL_ACTIONS = new Set(["tmux-bridge-url", "tmux-url", "bridge-url"]);
const PYTHON_BRIDGE_URL_ACTIONS = new Set(["python-bridge-url", "python-url", "libreoffice-bridge-url"]);

const URL_CLEAR_ACTIONS = new Set(["clear", "unset", "none"]);
const URL_SHOW_ACTIONS = new Set(["show", "status", "get"]);

const TMUX_BRIDGE_TOKEN_ACTIONS = new Set(["tmux-bridge-token", "tmux-token", "bridge-token"]);
const PYTHON_BRIDGE_TOKEN_ACTIONS = new Set(["python-bridge-token", "python-token", "libreoffice-bridge-token"]);

const TOKEN_CLEAR_ACTIONS = new Set(["clear", "unset", "none"]);
const TOKEN_SHOW_ACTIONS = new Set(["show", "status", "get"]);
const TMUX_STATUS_ACTIONS = new Set(["tmux-status", "tmux-bridge-status", "bridge-status"]);
const TMUX_BRIDGE_HEALTH_TIMEOUT_MS = 1500;

type FeatureResolver = (input: string) => ExperimentalFeatureDefinition | null;

export interface TmuxBridgeHealthStatus {
  reachable: boolean;
  status?: number;
  mode?: string;
  backend?: string;
  sessions?: number;
  error?: string;
}

export interface ExperimentalCommandDependencies {
  showToast?: (message: string) => void;
  showExperimentalDialog?: () => void;
  getFeatureSlugs?: () => string[];
  resolveFeature?: FeatureResolver;
  setFeatureEnabled?: (featureId: ExperimentalFeatureId, enabled: boolean) => void;
  toggleFeature?: (featureId: ExperimentalFeatureId) => boolean;

  getTmuxBridgeUrl?: () => Promise<string | undefined>;
  setTmuxBridgeUrl?: (url: string) => Promise<void>;
  clearTmuxBridgeUrl?: () => Promise<void>;
  validateTmuxBridgeUrl?: (url: string) => string;

  getTmuxBridgeToken?: () => Promise<string | undefined>;
  setTmuxBridgeToken?: (token: string) => Promise<void>;
  clearTmuxBridgeToken?: () => Promise<void>;
  validateTmuxBridgeToken?: (token: string) => string;
  isTmuxBridgeEnabled?: () => boolean;
  probeTmuxBridgeHealth?: (bridgeUrl: string) => Promise<TmuxBridgeHealthStatus>;

  getPythonBridgeUrl?: () => Promise<string | undefined>;
  setPythonBridgeUrl?: (url: string) => Promise<void>;
  clearPythonBridgeUrl?: () => Promise<void>;
  validatePythonBridgeUrl?: (url: string) => string;

  getPythonBridgeToken?: () => Promise<string | undefined>;
  setPythonBridgeToken?: (token: string) => Promise<void>;
  clearPythonBridgeToken?: () => Promise<void>;
  validatePythonBridgeToken?: (token: string) => string;

  notifyToolConfigChanged?: (configKey: string) => void;
}

interface ResolvedExperimentalCommandDependencies {
  showToast: (message: string) => void;
  showExperimentalDialog: () => void;
  getFeatureSlugs: () => string[];
  resolveFeature: FeatureResolver;
  setFeatureEnabled: (featureId: ExperimentalFeatureId, enabled: boolean) => void;
  toggleFeature: (featureId: ExperimentalFeatureId) => boolean;

  getTmuxBridgeUrl: () => Promise<string | undefined>;
  setTmuxBridgeUrl: (url: string) => Promise<void>;
  clearTmuxBridgeUrl: () => Promise<void>;
  validateTmuxBridgeUrl: (url: string) => string;

  getTmuxBridgeToken: () => Promise<string | undefined>;
  setTmuxBridgeToken: (token: string) => Promise<void>;
  clearTmuxBridgeToken: () => Promise<void>;
  validateTmuxBridgeToken: (token: string) => string;
  isTmuxBridgeEnabled: () => boolean;
  probeTmuxBridgeHealth: (bridgeUrl: string) => Promise<TmuxBridgeHealthStatus>;

  getPythonBridgeUrl: () => Promise<string | undefined>;
  setPythonBridgeUrl: (url: string) => Promise<void>;
  clearPythonBridgeUrl: () => Promise<void>;
  validatePythonBridgeUrl: (url: string) => string;

  getPythonBridgeToken: () => Promise<string | undefined>;
  setPythonBridgeToken: (token: string) => Promise<void>;
  clearPythonBridgeToken: () => Promise<void>;
  validatePythonBridgeToken: (token: string) => string;

  notifyToolConfigChanged: (configKey: string) => void;
}

interface BridgeUrlCommandConfig {
  bridgeLabel: string;
  commandLabel: string;
  exampleUrl: string;
  configKey: string;
  getValue: () => Promise<string | undefined>;
  setValue: (url: string) => Promise<void>;
  clearValue: () => Promise<void>;
  validate: (url: string) => string;
  showToast: (message: string) => void;
  notifyConfigChanged: (configKey: string) => void;
}

interface BridgeTokenCommandConfig {
  bridgeLabel: string;
  commandLabel: string;
  configKey: string;
  getValue: () => Promise<string | undefined>;
  setValue: (token: string) => Promise<void>;
  clearValue: () => Promise<void>;
  validate: (token: string) => string;
  showToast: (message: string) => void;
  notifyConfigChanged: (configKey: string) => void;
}

function tokenize(args: string): string[] {
  return args
    .trim()
    .split(/\s+/)
    .filter((token) => token.length > 0);
}

function normalizeFeatureToken(input: string): string {
  return input.trim().toLowerCase().replace(/[\s_]+/g, "-");
}

function getLegacyFeatureRedirectMessage(featureArg: string): string | null {
  const normalized = normalizeFeatureToken(featureArg);
  if (!LEGACY_INTEGRATIONS_FEATURES.has(normalized)) {
    return null;
  }

  return "External tools (including MCP) are managed in /integrations (Tools & MCP), not /experimental.";
}

function usageText(): string {
  return (
    "Usage: /experimental [list|on|off|toggle] <feature> " +
    "| /experimental tmux-bridge-url [<url>|show|clear] " +
    "| /experimental tmux-bridge-token [<token>|show|clear] " +
    "| /experimental tmux-status " +
    "| /experimental python-bridge-url [<url>|show|clear] " +
    "| /experimental python-bridge-token [<token>|show|clear]"
  );
}

function featureListText(getFeatureSlugs: () => string[]): string {
  const slugs = getFeatureSlugs();
  return slugs.length > 0 ? slugs.join(", ") : "(none)";
}

async function getSettingsStore() {
  const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
  return storageModule.getAppStorage().settings;
}

async function defaultGetBridgeUrl(settingKey: string): Promise<string | undefined> {
  try {
    const settings = await getSettingsStore();
    const value = await settings.get<string>(settingKey);
    if (typeof value !== "string") return undefined;

    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  } catch {
    return undefined;
  }
}

async function defaultSetSettingValue(settingKey: string, value: string): Promise<void> {
  const settings = await getSettingsStore();
  await settings.set(settingKey, value);
}

async function defaultClearSettingValue(settingKey: string): Promise<void> {
  const settings = await getSettingsStore();
  await settings.delete(settingKey);
}

function defaultValidateBridgeToken(label: string, token: string): string {
  const normalized = token.trim();
  if (normalized.length === 0) {
    throw new Error(`${label} cannot be empty.`);
  }

  if (/\s/u.test(normalized)) {
    throw new Error(`${label} must not contain whitespace.`);
  }

  if (normalized.length > 512) {
    throw new Error(`${label} is too long (max 512 characters).`);
  }

  return normalized;
}

function defaultValidateTmuxBridgeToken(token: string): string {
  return defaultValidateBridgeToken("Tmux bridge token", token);
}

function defaultValidatePythonBridgeToken(token: string): string {
  return defaultValidateBridgeToken("Python bridge token", token);
}

function maskToken(token: string): string {
  if (token.length <= 4) {
    return "*".repeat(token.length);
  }

  if (token.length <= 8) {
    return `${token.slice(0, 2)}${"*".repeat(token.length - 2)}`;
  }

  const hiddenLength = token.length - 6;
  return `${token.slice(0, 4)}${"*".repeat(hiddenLength)}${token.slice(-2)}`;
}

function defaultIsTmuxBridgeEnabled(): boolean {
  return isExperimentalFeatureEnabled("tmux_bridge");
}

function normalizeOptionalString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function normalizeOptionalInteger(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isInteger(value)) {
    return undefined;
  }

  return value;
}

function tryParseJson(text: string): unknown {
  const trimmed = text.trim();
  if (trimmed.length === 0) return null;

  try {
    return JSON.parse(trimmed);
  } catch {
    return null;
  }
}

function resolveDependencies(
  dependencies: ExperimentalCommandDependencies,
): ResolvedExperimentalCommandDependencies {
  return {
    showToast: dependencies.showToast ?? showToast,
    showExperimentalDialog: dependencies.showExperimentalDialog ?? showExperimentalDialog,
    getFeatureSlugs: dependencies.getFeatureSlugs ?? getExperimentalFeatureSlugs,
    resolveFeature: dependencies.resolveFeature ?? resolveExperimentalFeature,
    setFeatureEnabled: dependencies.setFeatureEnabled ?? setExperimentalFeatureEnabled,
    toggleFeature: dependencies.toggleFeature ?? toggleExperimentalFeature,

    getTmuxBridgeUrl:
      dependencies.getTmuxBridgeUrl
      ?? (() => defaultGetBridgeUrl(TMUX_BRIDGE_URL_SETTING_KEY)),
    setTmuxBridgeUrl:
      dependencies.setTmuxBridgeUrl
      ?? ((url: string) => defaultSetSettingValue(TMUX_BRIDGE_URL_SETTING_KEY, url)),
    clearTmuxBridgeUrl:
      dependencies.clearTmuxBridgeUrl
      ?? (() => defaultClearSettingValue(TMUX_BRIDGE_URL_SETTING_KEY)),
    validateTmuxBridgeUrl: dependencies.validateTmuxBridgeUrl ?? validateOfficeProxyUrl,

    getTmuxBridgeToken:
      dependencies.getTmuxBridgeToken
      ?? (() => defaultGetBridgeUrl(TMUX_BRIDGE_TOKEN_SETTING_KEY)),
    setTmuxBridgeToken:
      dependencies.setTmuxBridgeToken
      ?? ((token: string) => defaultSetSettingValue(TMUX_BRIDGE_TOKEN_SETTING_KEY, token)),
    clearTmuxBridgeToken:
      dependencies.clearTmuxBridgeToken
      ?? (() => defaultClearSettingValue(TMUX_BRIDGE_TOKEN_SETTING_KEY)),
    validateTmuxBridgeToken: dependencies.validateTmuxBridgeToken ?? defaultValidateTmuxBridgeToken,
    isTmuxBridgeEnabled: dependencies.isTmuxBridgeEnabled ?? defaultIsTmuxBridgeEnabled,
    probeTmuxBridgeHealth: dependencies.probeTmuxBridgeHealth ?? defaultProbeTmuxBridgeHealth,

    getPythonBridgeUrl:
      dependencies.getPythonBridgeUrl
      ?? (() => defaultGetBridgeUrl(PYTHON_BRIDGE_URL_SETTING_KEY)),
    setPythonBridgeUrl:
      dependencies.setPythonBridgeUrl
      ?? ((url: string) => defaultSetSettingValue(PYTHON_BRIDGE_URL_SETTING_KEY, url)),
    clearPythonBridgeUrl:
      dependencies.clearPythonBridgeUrl
      ?? (() => defaultClearSettingValue(PYTHON_BRIDGE_URL_SETTING_KEY)),
    validatePythonBridgeUrl: dependencies.validatePythonBridgeUrl ?? validateOfficeProxyUrl,

    getPythonBridgeToken:
      dependencies.getPythonBridgeToken
      ?? (() => defaultGetBridgeUrl(PYTHON_BRIDGE_TOKEN_SETTING_KEY)),
    setPythonBridgeToken:
      dependencies.setPythonBridgeToken
      ?? ((token: string) => defaultSetSettingValue(PYTHON_BRIDGE_TOKEN_SETTING_KEY, token)),
    clearPythonBridgeToken:
      dependencies.clearPythonBridgeToken
      ?? (() => defaultClearSettingValue(PYTHON_BRIDGE_TOKEN_SETTING_KEY)),
    validatePythonBridgeToken: dependencies.validatePythonBridgeToken ?? defaultValidatePythonBridgeToken,

    notifyToolConfigChanged: dependencies.notifyToolConfigChanged ?? ((configKey: string) => {
      dispatchExperimentalToolConfigChanged({ configKey });
    }),
  };
}

function asErrorMessage(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }
  return fallback;
}

async function defaultProbeTmuxBridgeHealth(bridgeUrl: string): Promise<TmuxBridgeHealthStatus> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => {
    controller.abort();
  }, TMUX_BRIDGE_HEALTH_TIMEOUT_MS);

  try {
    const target = `${bridgeUrl.replace(/\/+$/u, "")}/health`;
    const response = await fetch(target, {
      method: "GET",
      signal: controller.signal,
    });

    const bodyText = await response.text();
    const parsed = tryParseJson(bodyText);

    const status = normalizeOptionalInteger(response.status);
    let mode: string | undefined;
    let backend: string | undefined;
    let sessions: number | undefined;
    let error: string | undefined;

    if (isRecord(parsed)) {
      mode = normalizeOptionalString(parsed.mode);
      backend = normalizeOptionalString(parsed.backend);
      sessions = normalizeOptionalInteger(parsed.sessions);
      error = normalizeOptionalString(parsed.error);
    } else if (!response.ok) {
      error = normalizeOptionalString(bodyText);
    }

    return {
      reachable: response.ok,
      status,
      mode,
      backend,
      sessions,
      error,
    };
  } catch (error: unknown) {
    return {
      reachable: false,
      error: asErrorMessage(error, "Health check failed."),
    };
  } finally {
    clearTimeout(timeoutId);
  }
}

async function handleBridgeUrlCommand(
  valueTokens: string[],
  config: BridgeUrlCommandConfig,
): Promise<void> {
  const {
    bridgeLabel,
    commandLabel,
    exampleUrl,
    configKey,
    getValue,
    setValue,
    clearValue,
    validate,
    showToast,
    notifyConfigChanged,
  } = config;

  if (valueTokens.length === 0) {
    const existing = await getValue();
    if (!existing) {
      showToast(`${bridgeLabel} URL is not set. Example: /experimental ${commandLabel} ${exampleUrl}`);
      return;
    }

    showToast(`${bridgeLabel} URL: ${existing}`);
    return;
  }

  const firstToken = valueTokens[0].toLowerCase();
  if (URL_SHOW_ACTIONS.has(firstToken)) {
    const existing = await getValue();
    if (!existing) {
      showToast(`${bridgeLabel} URL is not set. Example: /experimental ${commandLabel} ${exampleUrl}`);
      return;
    }

    showToast(`${bridgeLabel} URL: ${existing}`);
    return;
  }

  if (valueTokens.length === 1 && URL_CLEAR_ACTIONS.has(firstToken)) {
    await clearValue();
    notifyConfigChanged(configKey);
    showToast(`${bridgeLabel} URL cleared.`);
    return;
  }

  const candidateUrl = valueTokens.join(" ");
  const normalized = validate(candidateUrl);
  await setValue(normalized);
  notifyConfigChanged(configKey);
  showToast(`${bridgeLabel} URL set to ${normalized}`);
}

async function handleBridgeTokenCommand(
  valueTokens: string[],
  config: BridgeTokenCommandConfig,
): Promise<void> {
  const {
    bridgeLabel,
    commandLabel,
    configKey,
    getValue,
    setValue,
    clearValue,
    validate,
    showToast,
    notifyConfigChanged,
  } = config;

  if (valueTokens.length === 0) {
    const existing = await getValue();
    if (!existing) {
      showToast(`${bridgeLabel} token is not set. Example: /experimental ${commandLabel} <token>`);
      return;
    }

    showToast(`${bridgeLabel} token: ${maskToken(existing)} (length ${existing.length})`);
    return;
  }

  const firstToken = valueTokens[0].toLowerCase();
  if (TOKEN_SHOW_ACTIONS.has(firstToken)) {
    const existing = await getValue();
    if (!existing) {
      showToast(`${bridgeLabel} token is not set. Example: /experimental ${commandLabel} <token>`);
      return;
    }

    showToast(`${bridgeLabel} token: ${maskToken(existing)} (length ${existing.length})`);
    return;
  }

  if (valueTokens.length === 1 && TOKEN_CLEAR_ACTIONS.has(firstToken)) {
    await clearValue();
    notifyConfigChanged(configKey);
    showToast(`${bridgeLabel} token cleared.`);
    return;
  }

  const candidateToken = valueTokens.join(" ");
  const normalized = validate(candidateToken);
  await setValue(normalized);
  notifyConfigChanged(configKey);
  showToast(`${bridgeLabel} token set (${maskToken(normalized)}).`);
}

async function handleTmuxStatusCommand(
  dependencies: ResolvedExperimentalCommandDependencies,
): Promise<void> {
  const featureEnabled = dependencies.isTmuxBridgeEnabled();
  const configuredBridgeUrl = await dependencies.getTmuxBridgeUrl();
  const configuredToken = await dependencies.getTmuxBridgeToken();

  let normalizedBridgeUrl: string | undefined;
  let bridgeUrlValidationError: string | undefined;

  if (configuredBridgeUrl) {
    try {
      normalizedBridgeUrl = dependencies.validateTmuxBridgeUrl(configuredBridgeUrl);
    } catch (error: unknown) {
      bridgeUrlValidationError = asErrorMessage(error, "invalid bridge URL");
    }
  }

  const health = normalizedBridgeUrl
    ? await dependencies.probeTmuxBridgeHealth(normalizedBridgeUrl)
    : undefined;

  let gateReason: TmuxBridgeGateReason | undefined;
  if (!featureEnabled) {
    gateReason = "tmux_experiment_disabled";
  } else if (!configuredBridgeUrl) {
    gateReason = "missing_bridge_url";
  } else if (!normalizedBridgeUrl) {
    gateReason = "invalid_bridge_url";
  } else if (!health?.reachable) {
    gateReason = "bridge_unreachable";
  }

  const gate: TmuxBridgeGateResult = gateReason
    ? {
      allowed: false,
      reason: gateReason,
      bridgeUrl: normalizedBridgeUrl,
    }
    : {
      allowed: true,
      bridgeUrl: normalizedBridgeUrl,
    };

  const lines: string[] = ["Tmux bridge status:"];
  lines.push(`- feature flag (tmux-bridge): ${featureEnabled ? "enabled" : "disabled"}`);

  if (!configuredBridgeUrl) {
    lines.push("- bridge URL: not set");
  } else if (normalizedBridgeUrl) {
    lines.push(`- bridge URL: ${normalizedBridgeUrl}`);
  } else {
    lines.push(`- bridge URL: invalid (${bridgeUrlValidationError ?? configuredBridgeUrl})`);
  }

  if (!configuredToken) {
    lines.push("- auth token: not set");
  } else {
    lines.push(`- auth token: set (${maskToken(configuredToken)}, length ${configuredToken.length})`);
  }

  if (gate.allowed) {
    lines.push("- gate: pass");
  } else {
    const reason = gate.reason ?? "bridge_unreachable";
    lines.push(`- gate: blocked (${reason})`);
    lines.push(`  hint: ${buildTmuxBridgeGateErrorMessage(reason)}`);
  }

  if (!health) {
    lines.push("- health: not checked (set a valid tmux bridge URL first)");
  } else if (health.reachable) {
    const details: string[] = [];
    if (health.status !== undefined) details.push(`HTTP ${health.status}`);
    if (health.mode) details.push(`mode=${health.mode}`);
    if (health.backend) details.push(`backend=${health.backend}`);
    if (health.sessions !== undefined) details.push(`sessions=${health.sessions}`);

    const suffix = details.length > 0 ? ` (${details.join(", ")})` : "";
    lines.push(`- health: reachable${suffix}`);
  } else {
    const details: string[] = [];
    if (health.status !== undefined) details.push(`HTTP ${health.status}`);
    if (health.error) details.push(health.error);

    const suffix = details.length > 0 ? ` (${details.join("; ")})` : "";
    lines.push(`- health: unreachable${suffix}`);
  }

  dependencies.showToast(lines.join("\n"));
}

export function createExperimentalCommands(
  dependencies: ExperimentalCommandDependencies = {},
): SlashCommand[] {
  const resolved = resolveDependencies(dependencies);

  return [
    {
      name: "experimental",
      description: "Manage experimental features",
      source: "builtin",
      execute: async (args: string) => {
        try {
          const tokens = tokenize(args);
          if (tokens.length === 0) {
            resolved.showExperimentalDialog();
            return;
          }

          const action = tokens[0].toLowerCase();

          if (action === "help") {
            resolved.showToast(`${usageText()} • Features: ${featureListText(resolved.getFeatureSlugs)}`);
            return;
          }

          if (OPEN_ACTIONS.has(action)) {
            resolved.showExperimentalDialog();
            return;
          }

          if (TMUX_BRIDGE_URL_ACTIONS.has(action)) {
            await handleBridgeUrlCommand(tokens.slice(1), {
              bridgeLabel: "Tmux bridge",
              commandLabel: "tmux-bridge-url",
              exampleUrl: "https://localhost:3337",
              configKey: TMUX_BRIDGE_URL_SETTING_KEY,
              getValue: resolved.getTmuxBridgeUrl,
              setValue: resolved.setTmuxBridgeUrl,
              clearValue: resolved.clearTmuxBridgeUrl,
              validate: resolved.validateTmuxBridgeUrl,
              showToast: resolved.showToast,
              notifyConfigChanged: resolved.notifyToolConfigChanged,
            });
            return;
          }

          if (TMUX_BRIDGE_TOKEN_ACTIONS.has(action)) {
            await handleBridgeTokenCommand(tokens.slice(1), {
              bridgeLabel: "Tmux bridge",
              commandLabel: "tmux-bridge-token",
              configKey: TMUX_BRIDGE_TOKEN_SETTING_KEY,
              getValue: resolved.getTmuxBridgeToken,
              setValue: resolved.setTmuxBridgeToken,
              clearValue: resolved.clearTmuxBridgeToken,
              validate: resolved.validateTmuxBridgeToken,
              showToast: resolved.showToast,
              notifyConfigChanged: resolved.notifyToolConfigChanged,
            });
            return;
          }

          if (TMUX_STATUS_ACTIONS.has(action)) {
            if (tokens.length > 1) {
              resolved.showToast("Usage: /experimental tmux-status");
              return;
            }

            await handleTmuxStatusCommand(resolved);
            return;
          }

          if (PYTHON_BRIDGE_URL_ACTIONS.has(action)) {
            await handleBridgeUrlCommand(tokens.slice(1), {
              bridgeLabel: "Python bridge",
              commandLabel: "python-bridge-url",
              exampleUrl: "https://localhost:3340",
              configKey: PYTHON_BRIDGE_URL_SETTING_KEY,
              getValue: resolved.getPythonBridgeUrl,
              setValue: resolved.setPythonBridgeUrl,
              clearValue: resolved.clearPythonBridgeUrl,
              validate: resolved.validatePythonBridgeUrl,
              showToast: resolved.showToast,
              notifyConfigChanged: resolved.notifyToolConfigChanged,
            });
            return;
          }

          if (PYTHON_BRIDGE_TOKEN_ACTIONS.has(action)) {
            await handleBridgeTokenCommand(tokens.slice(1), {
              bridgeLabel: "Python bridge",
              commandLabel: "python-bridge-token",
              configKey: PYTHON_BRIDGE_TOKEN_SETTING_KEY,
              getValue: resolved.getPythonBridgeToken,
              setValue: resolved.setPythonBridgeToken,
              clearValue: resolved.clearPythonBridgeToken,
              validate: resolved.validatePythonBridgeToken,
              showToast: resolved.showToast,
              notifyConfigChanged: resolved.notifyToolConfigChanged,
            });
            return;
          }

          const isToggleAction =
            ENABLE_ACTIONS.has(action)
            || DISABLE_ACTIONS.has(action)
            || TOGGLE_ACTIONS.has(action);

          if (!isToggleAction) {
            resolved.showToast(usageText());
            return;
          }

          const featureArg = tokens.slice(1).join(" ");
          if (!featureArg) {
            resolved.showToast(`${usageText()} • Features: ${featureListText(resolved.getFeatureSlugs)}`);
            return;
          }

          const feature = resolved.resolveFeature(featureArg);
          if (!feature) {
            const redirectMessage = getLegacyFeatureRedirectMessage(featureArg);
            if (redirectMessage) {
              resolved.showToast(redirectMessage);
              return;
            }

            resolved.showToast(
              `Unknown feature: ${featureArg}. Available: ${featureListText(resolved.getFeatureSlugs)}`,
            );
            return;
          }

          let enabled = false;

          if (ENABLE_ACTIONS.has(action)) {
            resolved.setFeatureEnabled(feature.id, true);
            enabled = true;
          } else if (DISABLE_ACTIONS.has(action)) {
            resolved.setFeatureEnabled(feature.id, false);
            enabled = false;
          } else {
            enabled = resolved.toggleFeature(feature.id);
          }

          const suffix = feature.wiring === "flag-only"
            ? " (flag saved; feature not wired yet)"
            : "";
          resolved.showToast(`${feature.title}: ${enabled ? "enabled" : "disabled"}${suffix}`);
        } catch (error: unknown) {
          resolved.showToast(asErrorMessage(error, "Failed to run /experimental command."));
        }
      },
    },
  ];
}
