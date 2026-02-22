import { validateOfficeProxyUrl } from "../auth/proxy-validation.js";

import { joinBridgeUrl } from "./bridge-http-utils.js";

const BRIDGE_HEALTH_PATH = "/health";
const BRIDGE_HEALTH_TIMEOUT_MS = 900;

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

export async function probeBridgeHealth(bridgeUrl: string): Promise<boolean> {
  const response = await fetchBridgeHealthResponse(bridgeUrl);
  return response?.ok === true;
}

export async function fetchBridgeHealthJson(bridgeUrl: string): Promise<unknown> {
  const response = await fetchBridgeHealthResponse(bridgeUrl);
  if (!response?.ok) {
    return null;
  }

  try {
    return await response.json() as unknown;
  } catch {
    return null;
  }
}

export async function getBridgeSetting(settingKey: string): Promise<string | undefined> {
  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
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
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const storage = storageModule.getAppStorage();
    await storage.settings.set(settingKey, value);
  } catch {
    // ignore (approval prompt will continue to appear if persistence is unavailable)
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
