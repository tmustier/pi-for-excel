function isAuthOauthStoragePayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * OAuth credential persistence for in-browser OAuth flows.
 *
 * Credentials are stored in IndexedDB via SettingsStore under `oauth.<providerId>`.
 * We intentionally do not read/write legacy localStorage keys anymore.
 *
 * Security note: this is storage hygiene, not an XSS boundary. Same-origin script
 * execution can read IndexedDB.
 */

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

/**
 * Load OAuth credentials from IndexedDB settings.
 */
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
