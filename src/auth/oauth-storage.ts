function isAuthOauthStoragePayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

import type { OAuthCredentials } from "@earendil-works/pi-ai";
import type { SettingsAccess } from "../storage/local/settings-store.js";
export function isOAuthCredentials(value: unknown): value is OAuthCredentials {
  return (
    isAuthOauthStoragePayloadShape(value) &&
    typeof value.refresh === "string" &&
    typeof value.access === "string" &&
    typeof value.expires === "number"
  );
}

function oauthSettingsKey(providerId: string): string {
  return `oauth.${providerId}`;
}

export async function loadOAuthCredentials(
  settings: SettingsAccess,
  providerId: string,
): Promise<OAuthCredentials | null> {
  try {
    const stored: unknown = await settings.get(oauthSettingsKey(providerId));
    return isOAuthCredentials(stored) ? stored : null;
  } catch {
    return null;
  }
}

export async function saveOAuthCredentials(
  settings: SettingsAccess,
  providerId: string,
  credentials: OAuthCredentials,
): Promise<void> {
  await settings.set(oauthSettingsKey(providerId), credentials);
}

export async function clearOAuthCredentials(
  settings: SettingsAccess,
  providerId: string,
): Promise<void> {
  await settings.delete(oauthSettingsKey(providerId));
}
