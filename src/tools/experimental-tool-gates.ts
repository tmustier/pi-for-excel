/**
 * Experimental tool gatekeeper.
 *
 * Security posture for experimental capabilities (local bridges + files workspace):
 * - capability must be explicitly enabled via /experimental
 * - local bridge URL must be configured (for bridge-backed tools)
 * - bridge must be reachable at execution time (for bridge-backed tools)
 * - tools remain registered (stable tool list / prompt caching)
 * - execution performs a hard gate check (defense in depth)
 */

import type { AgentTool } from "@mariozechner/pi-agent-core";

import { validateOfficeProxyUrl } from "../auth/proxy-validation.js";
import { isExperimentalFeatureEnabled } from "../experiments/flags.js";

const TMUX_TOOL_NAME = "tmux";
const FILES_TOOL_NAME = "files";
const PYTHON_TOOL_NAMES = new Set<string>([
  "python_run",
  "libreoffice_convert",
  "python_transform_range",
]);

const BRIDGE_HEALTH_PATH = "/health";
const BRIDGE_HEALTH_TIMEOUT_MS = 900;

export const TMUX_BRIDGE_URL_SETTING_KEY = "tmux.bridge.url";
export const PYTHON_BRIDGE_URL_SETTING_KEY = "python.bridge.url";
export const PYTHON_BRIDGE_APPROVED_URL_SETTING_KEY = "python.bridge.approved.url";

export type TmuxBridgeGateReason =
  | "tmux_experiment_disabled"
  | "missing_bridge_url"
  | "invalid_bridge_url"
  | "bridge_unreachable";

export interface TmuxBridgeGateResult {
  allowed: boolean;
  bridgeUrl?: string;
  reason?: TmuxBridgeGateReason;
}

export interface TmuxBridgeGateDependencies {
  isTmuxExperimentEnabled?: () => boolean;
  getTmuxBridgeUrl?: () => Promise<string | undefined>;
  validateBridgeUrl?: (url: string) => string | null;
  probeTmuxBridge?: (bridgeUrl: string) => Promise<boolean>;
}

export type PythonBridgeGateReason =
  | "python_experiment_disabled"
  | "missing_bridge_url"
  | "invalid_bridge_url"
  | "bridge_unreachable";

export interface PythonBridgeGateResult {
  allowed: boolean;
  bridgeUrl?: string;
  reason?: PythonBridgeGateReason;
}

export interface PythonBridgeGateDependencies {
  isPythonExperimentEnabled?: () => boolean;
  getPythonBridgeUrl?: () => Promise<string | undefined>;
  validatePythonBridgeUrl?: (url: string) => string | null;
  probePythonBridge?: (bridgeUrl: string) => Promise<boolean>;
}

export type FilesWorkspaceGateReason = "files_experiment_disabled";

export interface FilesWorkspaceGateResult {
  allowed: boolean;
  reason?: FilesWorkspaceGateReason;
}

export interface FilesWorkspaceGateDependencies {
  isFilesWorkspaceExperimentEnabled?: () => boolean;
}

export interface PythonBridgeApprovalRequest {
  toolName: string;
  bridgeUrl: string;
  params: unknown;
}

export interface ExperimentalToolGateDependencies extends
  TmuxBridgeGateDependencies,
  PythonBridgeGateDependencies,
  FilesWorkspaceGateDependencies {
  requestPythonBridgeApproval?: (request: PythonBridgeApprovalRequest) => Promise<boolean>;
  getApprovedPythonBridgeUrl?: () => Promise<string | undefined>;
  setApprovedPythonBridgeUrl?: (bridgeUrl: string) => Promise<void>;
}

function defaultIsTmuxExperimentEnabled(): boolean {
  return isExperimentalFeatureEnabled("tmux_bridge");
}

function defaultIsPythonExperimentEnabled(): boolean {
  return isExperimentalFeatureEnabled("python_bridge");
}

function defaultIsFilesWorkspaceExperimentEnabled(): boolean {
  return isExperimentalFeatureEnabled("files_workspace");
}

async function defaultGetBridgeUrl(settingKey: string): Promise<string | undefined> {
  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const storage = storageModule.getAppStorage();
    const value = await storage.settings.get<string>(settingKey);
    if (typeof value !== "string") return undefined;

    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  } catch {
    return undefined;
  }
}

async function defaultGetTmuxBridgeUrl(): Promise<string | undefined> {
  return defaultGetBridgeUrl(TMUX_BRIDGE_URL_SETTING_KEY);
}

async function defaultGetPythonBridgeUrl(): Promise<string | undefined> {
  return defaultGetBridgeUrl(PYTHON_BRIDGE_URL_SETTING_KEY);
}

async function defaultSetBridgeSetting(settingKey: string, value: string): Promise<void> {
  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const storage = storageModule.getAppStorage();
    await storage.settings.set(settingKey, value);
  } catch {
    // ignore (approval prompt will continue to appear if persistence is unavailable)
  }
}

async function defaultGetApprovedPythonBridgeUrl(): Promise<string | undefined> {
  return defaultGetBridgeUrl(PYTHON_BRIDGE_APPROVED_URL_SETTING_KEY);
}

async function defaultSetApprovedPythonBridgeUrl(bridgeUrl: string): Promise<void> {
  await defaultSetBridgeSetting(PYTHON_BRIDGE_APPROVED_URL_SETTING_KEY, bridgeUrl);
}

function defaultValidateBridgeUrl(url: string): string | null {
  try {
    return validateOfficeProxyUrl(url);
  } catch {
    return null;
  }
}

async function defaultProbeBridge(bridgeUrl: string): Promise<boolean> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => {
    controller.abort();
  }, BRIDGE_HEALTH_TIMEOUT_MS);

  try {
    const target = `${bridgeUrl.replace(/\/+$/, "")}${BRIDGE_HEALTH_PATH}`;
    const response = await fetch(target, {
      method: "GET",
      signal: controller.signal,
    });
    return response.ok;
  } catch {
    return false;
  } finally {
    clearTimeout(timeoutId);
  }
}

export async function evaluateTmuxBridgeGate(
  dependencies: TmuxBridgeGateDependencies = {},
): Promise<TmuxBridgeGateResult> {
  const isEnabled = dependencies.isTmuxExperimentEnabled ?? defaultIsTmuxExperimentEnabled;
  if (!isEnabled()) {
    return {
      allowed: false,
      reason: "tmux_experiment_disabled",
    };
  }

  const getBridgeUrl = dependencies.getTmuxBridgeUrl ?? defaultGetTmuxBridgeUrl;
  const rawBridgeUrl = await getBridgeUrl();
  if (!rawBridgeUrl) {
    return {
      allowed: false,
      reason: "missing_bridge_url",
    };
  }

  const validateBridgeUrl = dependencies.validateBridgeUrl ?? defaultValidateBridgeUrl;
  const bridgeUrl = validateBridgeUrl(rawBridgeUrl);
  if (!bridgeUrl) {
    return {
      allowed: false,
      reason: "invalid_bridge_url",
    };
  }

  const probeTmuxBridge = dependencies.probeTmuxBridge ?? defaultProbeBridge;
  const reachable = await probeTmuxBridge(bridgeUrl);
  if (!reachable) {
    return {
      allowed: false,
      reason: "bridge_unreachable",
      bridgeUrl,
    };
  }

  return {
    allowed: true,
    bridgeUrl,
  };
}

export async function evaluatePythonBridgeGate(
  dependencies: PythonBridgeGateDependencies = {},
): Promise<PythonBridgeGateResult> {
  const isEnabled = dependencies.isPythonExperimentEnabled ?? defaultIsPythonExperimentEnabled;
  if (!isEnabled()) {
    return {
      allowed: false,
      reason: "python_experiment_disabled",
    };
  }

  const getBridgeUrl = dependencies.getPythonBridgeUrl ?? defaultGetPythonBridgeUrl;
  const rawBridgeUrl = await getBridgeUrl();
  if (!rawBridgeUrl) {
    return {
      allowed: false,
      reason: "missing_bridge_url",
    };
  }

  const validateBridgeUrl = dependencies.validatePythonBridgeUrl ?? defaultValidateBridgeUrl;
  const bridgeUrl = validateBridgeUrl(rawBridgeUrl);
  if (!bridgeUrl) {
    return {
      allowed: false,
      reason: "invalid_bridge_url",
    };
  }

  const probePythonBridge = dependencies.probePythonBridge ?? defaultProbeBridge;
  const reachable = await probePythonBridge(bridgeUrl);
  if (!reachable) {
    return {
      allowed: false,
      reason: "bridge_unreachable",
      bridgeUrl,
    };
  }

  return {
    allowed: true,
    bridgeUrl,
  };
}

export function evaluateFilesWorkspaceGate(
  dependencies: FilesWorkspaceGateDependencies = {},
): FilesWorkspaceGateResult {
  const isEnabled =
    dependencies.isFilesWorkspaceExperimentEnabled
    ?? defaultIsFilesWorkspaceExperimentEnabled;

  if (!isEnabled()) {
    return {
      allowed: false,
      reason: "files_experiment_disabled",
    };
  }

  return { allowed: true };
}

export function buildTmuxBridgeGateErrorMessage(reason: TmuxBridgeGateReason): string {
  switch (reason) {
    case "tmux_experiment_disabled":
      return "Tmux bridge is disabled. Enable it with /experimental on tmux-bridge.";
    case "missing_bridge_url":
      return `Tmux bridge URL is not configured. Run /experimental tmux-bridge-url https://localhost:<port> (setting: ${TMUX_BRIDGE_URL_SETTING_KEY}).`;
    case "invalid_bridge_url":
      return "Tmux bridge URL is invalid. Use a full URL like https://localhost:3337.";
    case "bridge_unreachable":
      return "Tmux bridge is not reachable at the configured URL.";
  }
}

export function buildPythonBridgeGateErrorMessage(reason: PythonBridgeGateReason): string {
  switch (reason) {
    case "python_experiment_disabled":
      return "Python bridge is disabled. Enable it with /experimental on python-bridge.";
    case "missing_bridge_url":
      return `Python bridge URL is not configured. Run /experimental python-bridge-url https://localhost:<port> (setting: ${PYTHON_BRIDGE_URL_SETTING_KEY}).`;
    case "invalid_bridge_url":
      return "Python bridge URL is invalid. Use a full URL like https://localhost:3340.";
    case "bridge_unreachable":
      return "Python bridge is not reachable at the configured URL.";
  }
}

export function buildFilesWorkspaceGateErrorMessage(reason: FilesWorkspaceGateReason): string {
  switch (reason) {
    case "files_experiment_disabled":
      return "Files workspace is disabled. Enable it with /experimental on files-workspace.";
  }
}

function isRecordObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function getRecordValue(record: Record<string, unknown>, key: string): string | undefined {
  const value = record[key];
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}

function getPythonApprovalMessage(
  toolName: string,
  bridgeUrl: string,
  params: unknown,
): string {
  const title = "Allow local Python / LibreOffice execution?";

  if (isRecordObject(params)) {
    if (toolName === "python_run") {
      const code = getRecordValue(params, "code") ?? "(no code)";
      const previewLine = code.split("\n")[0] ?? code;
      return `${title}\n\nTool: python_run\nBridge: ${bridgeUrl}\nCode preview: ${previewLine}`;
    }

    if (toolName === "libreoffice_convert") {
      const inputPath = getRecordValue(params, "input_path") ?? "(unknown input)";
      const targetFormat = getRecordValue(params, "target_format") ?? "(unknown format)";
      return `${title}\n\nTool: libreoffice_convert\nBridge: ${bridgeUrl}\nInput: ${inputPath}\nTarget: ${targetFormat.toUpperCase()}`;
    }

    if (toolName === "python_transform_range") {
      const range = getRecordValue(params, "range") ?? "(unknown range)";
      const output = getRecordValue(params, "output_start_cell") ?? "(source top-left)";
      return `${title}\n\nTool: python_transform_range\nBridge: ${bridgeUrl}\nRange: ${range}\nOutput start: ${output}`;
    }
  }

  return `${title}\n\nTool: ${toolName}\nBridge: ${bridgeUrl}`;
}

function defaultRequestPythonBridgeApproval(
  request: PythonBridgeApprovalRequest,
): Promise<boolean> {
  if (typeof window === "undefined" || typeof window.confirm !== "function") {
    return Promise.resolve(true);
  }

  return Promise.resolve(
    window.confirm(getPythonApprovalMessage(request.toolName, request.bridgeUrl, request.params)),
  );
}

function wrapTmuxToolWithHardGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      const gate = await evaluateTmuxBridgeGate(dependencies);
      if (!gate.allowed) {
        const reason = gate.reason ?? "bridge_unreachable";
        throw new Error(buildTmuxBridgeGateErrorMessage(reason));
      }

      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

function wrapFilesToolWithHardGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      const gate = evaluateFilesWorkspaceGate(dependencies);
      if (!gate.allowed) {
        const reason = gate.reason ?? "files_experiment_disabled";
        throw new Error(buildFilesWorkspaceGateErrorMessage(reason));
      }

      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

function wrapPythonBridgeToolWithHardGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  const requestApproval = dependencies.requestPythonBridgeApproval ?? defaultRequestPythonBridgeApproval;
  const getApprovedBridgeUrl =
    dependencies.getApprovedPythonBridgeUrl
    ?? defaultGetApprovedPythonBridgeUrl;
  const setApprovedBridgeUrl =
    dependencies.setApprovedPythonBridgeUrl
    ?? defaultSetApprovedPythonBridgeUrl;

  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      const gate = await evaluatePythonBridgeGate(dependencies);
      if (!gate.allowed) {
        const reason = gate.reason ?? "bridge_unreachable";
        throw new Error(buildPythonBridgeGateErrorMessage(reason));
      }

      const bridgeUrl = gate.bridgeUrl;
      if (!bridgeUrl) {
        throw new Error("Python bridge gate did not return a bridge URL.");
      }

      const cachedApprovalUrl = await getApprovedBridgeUrl();
      if (cachedApprovalUrl !== bridgeUrl) {
        const approved = await requestApproval({
          toolName: tool.name,
          bridgeUrl,
          params,
        });
        if (!approved) {
          throw new Error("Python/LibreOffice execution cancelled by user.");
        }

        await setApprovedBridgeUrl(bridgeUrl);
      }

      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

/**
 * Apply experimental gates to tool execution.
 *
 * Current rules:
 * - `tmux`, `files`, `python_run`, `libreoffice_convert`, and `python_transform_range`
 *   stay registered to keep the tool list stable.
 * - each gated tool execution re-checks experiment flags (and bridge health where relevant).
 * - python/libreoffice bridge tools require user confirmation once per configured bridge URL.
 */
export function applyExperimentalToolGates(
  tools: AgentTool[],
  dependencies: ExperimentalToolGateDependencies = {},
): Promise<AgentTool[]> {
  const gatedTools: AgentTool[] = [];

  for (const tool of tools) {
    if (tool.name === TMUX_TOOL_NAME) {
      gatedTools.push(wrapTmuxToolWithHardGate(tool, dependencies));
      continue;
    }

    if (tool.name === FILES_TOOL_NAME) {
      gatedTools.push(wrapFilesToolWithHardGate(tool, dependencies));
      continue;
    }

    if (PYTHON_TOOL_NAMES.has(tool.name)) {
      gatedTools.push(wrapPythonBridgeToolWithHardGate(tool, dependencies));
      continue;
    }

    gatedTools.push(tool);
  }

  return Promise.resolve(gatedTools);
}
