import { validateOfficeProxyUrl } from "../auth/proxy-validation.js";

import { joinBridgeUrl } from "./bridge-http-utils.js";

const BRIDGE_HEALTH_PATH = "/health";
export const BRIDGE_HEALTH_TIMEOUT_MS = 900;

async function fetchBridgeHealthResponse(bridgeUrl: string): Promise<Response | null> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => {
    controller.abort();
  }, BRIDGE_HEALTH_TIMEOUT_MS);

  try {
    const response = await fetch(joinBridgeUrl(bridgeUrl, BRIDGE_HEALTH_PATH), {
      method: "GET",
      signal: controller.signal,
    });
    return response;
  } catch {
    return null;
  } finally {
    clearTimeout(timeoutId);
  }
}

function isBridgeHealthPayload(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export async function fetchBridgeHealthJson(bridgeUrl: string): Promise<DynamicValue> {
  const response = await fetchBridgeHealthResponse(bridgeUrl);
  if (!response?.ok) {
    return null;
  }

  try {
    return await response.json() as DynamicValue;
  } catch {
    return null;
  }
}

export async function probeTmuxBridgeHealth(bridgeUrl: string): Promise<boolean> {
  const payload = await fetchBridgeHealthJson(bridgeUrl);
  if (!isBridgeHealthPayload(payload) || payload.ok !== true) return false;
  return payload.mode === "tmux" || payload.backend === "tmux";
}

export type PythonBridgeCapability = "python" | "libreoffice";

export async function probePythonBridgeHealth(
  bridgeUrl: string,
  capability: PythonBridgeCapability = "python",
): Promise<boolean> {
  const payload = await fetchBridgeHealthJson(bridgeUrl);
  if (!isBridgeHealthPayload(payload) || payload.ok !== true) return false;
  if (payload.mode !== "real" && payload.backend !== "real") return false;

  const capabilityHealth = payload[capability];
  return isBridgeHealthPayload(capabilityHealth) && capabilityHealth.available === true;
}

export async function getBridgeSetting(settingKey: string): Promise<string | undefined> {
  try {
    const storageModule = await import("../storage/local/app-storage.js");
    const storage = storageModule.getAppStorage();
    const value = await storage.settings.get<string>(settingKey);
    if (typeof value !== "string") {
      return undefined;
    }

    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  } catch {
    return undefined;
  }
}

export async function setBridgeSetting(settingKey: string, value: string): Promise<void> {
  try {
    const storageModule = await import("../storage/local/app-storage.js");
    const storage = storageModule.getAppStorage();
    await storage.settings.set(settingKey, value);
  } catch {
    // Approval applies to this execution even if the optional cache cannot persist.
  }
}

export function validateBridgeUrl(url: string): string | null {
  try {
    return validateOfficeProxyUrl(url);
  } catch {
    return null;
  }
}

export function resolveValidatedBridgeUrl(
  configuredUrl: string | undefined,
  defaultUrl: string,
  validator: (url: string) => string | null = validateBridgeUrl,
): { bridgeUrl: string | null; usingDefaultBridgeUrl: boolean } {
  const usingDefaultBridgeUrl = !configuredUrl;
  const rawBridgeUrl = configuredUrl ?? defaultUrl;

  return {
    bridgeUrl: validator(rawBridgeUrl),
    usingDefaultBridgeUrl,
  };
}
