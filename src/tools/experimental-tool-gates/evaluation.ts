import {
  getBridgeSetting,
  probeBridgeHealth,
  resolveValidatedBridgeUrl,
  setBridgeSetting,
  validateBridgeUrl,
} from "../bridge-service-utils.js";

import {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
  PYTHON_BRIDGE_APPROVED_URL_SETTING_KEY,
  PYTHON_BRIDGE_URL_SETTING_KEY,
  TMUX_BRIDGE_URL_SETTING_KEY,
  type PythonBridgeGateDependencies,
  type PythonBridgeGateReason,
  type PythonBridgeGateResult,
  type TmuxBridgeGateDependencies,
  type TmuxBridgeGateReason,
  type TmuxBridgeGateResult,
} from "./types.js";

async function defaultGetTmuxBridgeUrl(): Promise<string | undefined> {
  return getBridgeSetting(TMUX_BRIDGE_URL_SETTING_KEY);
}

async function defaultGetPythonBridgeUrl(): Promise<string | undefined> {
  return getBridgeSetting(PYTHON_BRIDGE_URL_SETTING_KEY);
}

export async function defaultGetApprovedPythonBridgeUrl(): Promise<string | undefined> {
  return getBridgeSetting(PYTHON_BRIDGE_APPROVED_URL_SETTING_KEY);
}

export async function defaultSetApprovedPythonBridgeUrl(bridgeUrl: string): Promise<void> {
  await setBridgeSetting(PYTHON_BRIDGE_APPROVED_URL_SETTING_KEY, bridgeUrl);
}

const defaultValidateBridgeUrl = validateBridgeUrl;

const defaultProbeBridge = probeBridgeHealth;

export async function evaluateTmuxBridgeGate(
  dependencies: TmuxBridgeGateDependencies = {},
): Promise<TmuxBridgeGateResult> {
  // No experiment flag gate — tmux is available when bridge health passes.
  // If no URL override is configured, the default localhost bridge URL is probed.

  const getBridgeUrl = dependencies.getTmuxBridgeUrl ?? defaultGetTmuxBridgeUrl;
  const configuredBridgeUrl = await getBridgeUrl();

  const validateTmuxBridgeUrl = dependencies.validateBridgeUrl ?? defaultValidateBridgeUrl;
  const { bridgeUrl, usingDefaultBridgeUrl } = resolveValidatedBridgeUrl(
    configuredBridgeUrl,
    DEFAULT_TMUX_BRIDGE_URL,
    validateTmuxBridgeUrl,
  );
  if (!bridgeUrl) {
    return {
      allowed: false,
      reason: usingDefaultBridgeUrl ? "missing_bridge_url" : "invalid_bridge_url",
    };
  }

  const probeTmuxBridge = dependencies.probeTmuxBridge ?? defaultProbeBridge;
  const reachable = await probeTmuxBridge(bridgeUrl);
  if (!reachable) {
    return {
      allowed: false,
      reason: usingDefaultBridgeUrl ? "missing_bridge_url" : "bridge_unreachable",
      bridgeUrl,
    };
  }

  return {
    allowed: true,
    bridgeUrl,
  };
}

/**
 * Evaluate whether the Python/LibreOffice bridge is reachable.
 *
 * No experiment flag required. If no URL override is configured, the default
 * localhost bridge URL is probed automatically. The user approval dialog
 * (in wrappers.ts) serves as the security boundary.
 */
export async function evaluatePythonBridgeGate(
  dependencies: PythonBridgeGateDependencies = {},
): Promise<PythonBridgeGateResult> {
  const getBridgeUrl = dependencies.getPythonBridgeUrl ?? defaultGetPythonBridgeUrl;
  const configuredBridgeUrl = await getBridgeUrl();

  const validatePythonBridgeUrl = dependencies.validatePythonBridgeUrl ?? defaultValidateBridgeUrl;
  const { bridgeUrl, usingDefaultBridgeUrl } = resolveValidatedBridgeUrl(
    configuredBridgeUrl,
    DEFAULT_PYTHON_BRIDGE_URL,
    validatePythonBridgeUrl,
  );
  if (!bridgeUrl) {
    return {
      allowed: false,
      reason: usingDefaultBridgeUrl ? "missing_bridge_url" : "invalid_bridge_url",
    };
  }

  const probePythonBridge = dependencies.probePythonBridge ?? defaultProbeBridge;
  const reachable = await probePythonBridge(bridgeUrl);
  if (!reachable) {
    return {
      allowed: false,
      reason: usingDefaultBridgeUrl ? "missing_bridge_url" : "bridge_unreachable",
      bridgeUrl,
    };
  }

  return {
    allowed: true,
    bridgeUrl,
  };
}

export function buildTmuxBridgeGateErrorMessage(reason: TmuxBridgeGateReason): string {
  switch (reason) {
    case "missing_bridge_url":
      return (
        `Tmux bridge is not reachable at the default URL (${DEFAULT_TMUX_BRIDGE_URL}), ` +
        `and no URL override is configured (setting: ${TMUX_BRIDGE_URL_SETTING_KEY}).`
      );
    case "invalid_bridge_url":
      return "Tmux bridge URL is invalid. Use a full URL like https://localhost:3341.";
    case "bridge_unreachable":
      return "Tmux bridge is not reachable at the configured URL.";
  }
}

export function buildPythonBridgeGateErrorMessage(reason: PythonBridgeGateReason): string {
  switch (reason) {
    case "missing_bridge_url":
      return (
        `Python bridge is not reachable at the default URL (${DEFAULT_PYTHON_BRIDGE_URL}), ` +
        `and no URL override is configured (setting: ${PYTHON_BRIDGE_URL_SETTING_KEY}).`
      );
    case "invalid_bridge_url":
      return "Python bridge URL is invalid. Use a full URL like https://localhost:3340.";
    case "bridge_unreachable":
      return "Python bridge is not reachable at the configured URL.";
  }
}
