function isAuthRestorePayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Auto-restore auth credentials from pi's auth.json (dev) or browser storage (OAuth).
 *
 * Per-provider priority:
 * 1. pi's ~/.pi/agent/auth.json (served by Vite plugin at /__pi-auth)
 * 2. IndexedDB SettingsStore (`oauth.<providerId>`) as fallback
 */

import type { OAuthCredentials } from "@earendil-works/pi-ai";
import type { ProviderKeysStore } from "../storage/local/provider-keys-store.js";
import type { SettingsStore } from "../storage/local/settings-store.js";

import { originalFetch } from "./cors-proxy.js";
import { loadOAuthCredentials } from "./oauth-storage.js";
import { createOAuthRefresher } from "./oauth-refresh.js";
import type { BrowserOAuthProvider } from "./browser-oauth-types.js";
import { mapToApiProvider, BROWSER_OAUTH_PROVIDERS } from "./provider-map.js";
import { getOAuthProvider } from "./oauth-provider-registry.js";
import { getErrorMessage } from "../utils/errors.js";

type GetOAuthProvider = (id: string) => BrowserOAuthProvider | undefined;

type ApiKeyCredential = {
  type: "api_key";
  key: string;
};

type OAuthCredential = OAuthCredentials & {
  type: "oauth";
};

function isApiKeyCredential(value: unknown): value is ApiKeyCredential {
  return (
    isAuthRestorePayloadShape(value) &&
    value.type === "api_key" &&
    typeof value.key === "string" &&
    value.key.trim().length > 0
  );
}

function isOAuthCredential(value: unknown): value is OAuthCredential {
  return (
    isAuthRestorePayloadShape(value) &&
    value.type === "oauth" &&
    typeof value.refresh === "string" &&
    typeof value.access === "string" &&
    value.access.trim().length > 0 &&
    typeof value.expires === "number"
  );
}

/**
 * Restore credentials from all available sources.
 * Populates the ProviderKeysStore so ChatPanel can make API calls.
 */
export async function restoreCredentials(
  providerKeys: ProviderKeysStore,
  settings: SettingsStore,
  fetchAuth: typeof fetch = originalFetch,
): Promise<void> {
  const restoredProviders = await restoreFromPiAuth(providerKeys, getOAuthProvider, fetchAuth);
  await restoreFromBrowserOAuthStorage(providerKeys, settings, getOAuthProvider, restoredProviders);
}

async function restoreFromPiAuth(
  providerKeys: ProviderKeysStore,
  getOAuthProvider: GetOAuthProvider,
  fetchAuth: typeof fetch,
): Promise<Set<string>> {
  const restoredProviders = new Set<string>();

  try {
    const res = await fetchAuth("/__pi-auth");
    if (!res.ok) return restoredProviders;

    const authData: unknown = await res.json();
    if (!isAuthRestorePayloadShape(authData)) return restoredProviders;

    console.log(`[auth] Found pi auth.json with ${Object.keys(authData).length} provider(s)`);

    for (const [providerId, cred] of Object.entries(authData)) {
      try {
        const apiProvider = mapToApiProvider(providerId);

        if (isApiKeyCredential(cred)) {
          await providerKeys.set(apiProvider, cred.key);
          restoredProviders.add(apiProvider);
          console.log(`[auth] ${providerId}: API key loaded`);
          continue;
        }

        if (!isOAuthCredential(cred)) {
          continue;
        }

        const provider = getOAuthProvider(providerId);
        if (!provider) {
          console.log(`[auth] ${providerId}: no OAuth provider registered, skipping`);
          continue;
        }

        if (Date.now() >= cred.expires) {
          try {
            const refreshed = await provider.refreshToken(cred);
            await providerKeys.set(apiProvider, provider.getApiKey(refreshed));
            restoredProviders.add(apiProvider);
            console.log(`[auth] ${providerId}: token refreshed`);
          } catch (e) {
            console.warn(`[auth] ${providerId}: refresh failed (${getErrorMessage(e)})`);
          }
        } else {
          await providerKeys.set(apiProvider, provider.getApiKey(cred));
          restoredProviders.add(apiProvider);
          const hours = Math.round((cred.expires - Date.now()) / 3600000);
          console.log(`[auth] ${providerId}: OAuth token loaded (expires in ${hours}h)`);
        }
      } catch (e) {
        console.warn(`[auth] ${providerId}: failed (${getErrorMessage(e)})`);
      }
    }

    return restoredProviders;
  } catch {
    return restoredProviders;
  }
}

async function restoreFromBrowserOAuthStorage(
  providerKeys: ProviderKeysStore,
  settings: SettingsStore,
  getOAuthProvider: GetOAuthProvider,
  restoredProviders: Set<string>,
): Promise<void> {
  // Startup refresh goes through the same guarded path as on-demand refresh,
  // so two taskpanes opening together cannot both rotate one grant.
  const refresher = createOAuthRefresher({ settings, providerKeys, getOAuthProvider });

  for (const providerId of BROWSER_OAUTH_PROVIDERS) {
    const apiProvider = mapToApiProvider(providerId);
    if (restoredProviders.has(apiProvider)) continue;

    const credentials = await loadOAuthCredentials(settings, providerId);
    if (!credentials) continue;

    try {
      const provider = getOAuthProvider(providerId);
      if (!provider) continue;

      const outcome = await refresher.ensureFresh(apiProvider);
      if (outcome === "refreshed") {
        restoredProviders.add(apiProvider);
        console.log(`[auth] ${provider.name}: token refreshed from IndexedDB`);
      } else if (outcome === "fresh") {
        await providerKeys.set(apiProvider, provider.getApiKey(
          (await loadOAuthCredentials(settings, providerId)) ?? credentials,
        ));
        restoredProviders.add(apiProvider);
        console.log(`[auth] ${provider.name}: session restored from IndexedDB`);
      } else if (outcome === "failed") {
        // Refresh is attempted shortly before expiry; a token that still works is used until then.
        await providerKeys.set(apiProvider, provider.getApiKey(credentials));
        restoredProviders.add(apiProvider);
        console.warn(`[auth] ${provider.name}: refresh failed; it will be retried before the next request`);
      } else if (outcome === "expired") {
        console.warn(`[auth] ${provider.name}: refresh failed; it will be retried before the next request`);
      }
    } catch (e) {
      console.warn(`[auth] ${providerId}: failed to restore (${getErrorMessage(e)})`);
    }
  }
}
