/**
 * Auto-restore auth credentials from pi's auth.json (dev) or browser storage (OAuth).
 *
 * Priority:
 * 1. pi's ~/.pi/agent/auth.json (served by Vite plugin at /__pi-auth)
 * 2. IndexedDB SettingsStore (`oauth.<providerId>`)
 */

import type { OAuthCredentials, OAuthProviderInterface } from "@earendil-works/pi-ai";
import type { ProviderKeysStore } from "@earendil-works/pi-web-ui/dist/storage/stores/provider-keys-store.js";
import type { SettingsStore } from "@earendil-works/pi-web-ui/dist/storage/stores/settings-store.js";

import { originalFetch } from "./cors-proxy.js";
import { clearOAuthCredentials, loadOAuthCredentials, saveOAuthCredentials } from "./oauth-storage.js";
import { isOpenAICodexCredentialRefreshRequired } from "./openai-codex-browser-oauth.js";
import { mapToApiProvider, BROWSER_OAUTH_PROVIDERS } from "./provider-map.js";
import { getOAuthProvider } from "./oauth-provider-registry.js";
import { getErrorMessage } from "../utils/errors.js";
import { isRecord } from "../utils/type-guards.js";

type GetOAuthProvider = (id: string) => OAuthProviderInterface | undefined;

type ApiKeyCredential = {
  type: "api_key";
  key: string;
};

type OAuthCredential = OAuthCredentials & {
  type: "oauth";
};

function isApiKeyCredential(value: unknown): value is ApiKeyCredential {
  return (
    isRecord(value) &&
    value.type === "api_key" &&
    typeof value.key === "string" &&
    value.key.trim().length > 0
  );
}

function isOAuthCredential(value: unknown): value is OAuthCredential {
  return (
    isRecord(value) &&
    value.type === "oauth" &&
    typeof value.refresh === "string" &&
    typeof value.access === "string" &&
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
): Promise<void> {
  // 1. Try pi's auth.json (dev server only)
  if (await restoreFromPiAuth(providerKeys, getOAuthProvider)) {
    return;
  }

  // 2. Browser OAuth sessions (IndexedDB settings)
  await restoreFromBrowserOAuthStorage(providerKeys, settings, getOAuthProvider);
}

async function restoreFromPiAuth(
  providerKeys: ProviderKeysStore,
  getOAuthProvider: GetOAuthProvider,
): Promise<boolean> {
  try {
    const res = await originalFetch("/__pi-auth");
    if (!res.ok) return false;

    const authData: unknown = await res.json();
    if (!isRecord(authData)) return false;

    console.log(`[auth] Found pi auth.json with ${Object.keys(authData).length} provider(s)`);

    for (const [providerId, cred] of Object.entries(authData)) {
      try {
        const apiProvider = mapToApiProvider(providerId);

        if (isApiKeyCredential(cred)) {
          await providerKeys.set(apiProvider, cred.key);
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
            console.log(`[auth] ${providerId}: token refreshed`);
          } catch (e: unknown) {
            console.warn(`[auth] ${providerId}: refresh failed (${getErrorMessage(e)})`);
          }
        } else {
          await providerKeys.set(apiProvider, provider.getApiKey(cred));
          const hours = Math.round((cred.expires - Date.now()) / 3600000);
          console.log(`[auth] ${providerId}: OAuth token loaded (expires in ${hours}h)`);
        }
      } catch (e: unknown) {
        console.warn(`[auth] ${providerId}: failed (${getErrorMessage(e)})`);
      }
    }

    return true;
  } catch {
    return false;
  }
}

async function clearBrowserOAuthProvider(
  providerKeys: ProviderKeysStore,
  settings: SettingsStore,
  providerId: string,
): Promise<void> {
  await clearOAuthCredentials(settings, providerId).catch(() => {});
  await providerKeys.delete(mapToApiProvider(providerId)).catch(() => {});
}

async function restoreFromBrowserOAuthStorage(
  providerKeys: ProviderKeysStore,
  settings: SettingsStore,
  getOAuthProvider: GetOAuthProvider,
): Promise<void> {
  for (const providerId of BROWSER_OAUTH_PROVIDERS) {
    const credentials = await loadOAuthCredentials(settings, providerId);
    if (!credentials) continue;

    try {
      const provider = getOAuthProvider(providerId);
      if (!provider) continue;

      const apiProvider = mapToApiProvider(providerId);

      if (Date.now() >= credentials.expires) {
        try {
          const refreshed = await provider.refreshToken(credentials);
          await saveOAuthCredentials(settings, providerId, refreshed);
          await providerKeys.set(apiProvider, provider.getApiKey(refreshed));
          console.log(`[auth] ${provider.name}: token refreshed from IndexedDB`);
        } catch (e: unknown) {
          if (providerId === "openai-codex" && isOpenAICodexCredentialRefreshRequired(e)) {
            await clearBrowserOAuthProvider(providerKeys, settings, providerId);
            console.warn(`[auth] ${provider.name}: stored OAuth grant is stale; cleared credentials, please login again`);
          } else {
            console.warn(`[auth] ${provider.name}: refresh failed (${getErrorMessage(e)}), please login again`);
          }
        }
      } else {
        await providerKeys.set(apiProvider, provider.getApiKey(credentials));
        console.log(`[auth] ${provider.name}: session restored from IndexedDB`);
      }
    } catch (e: unknown) {
      if (providerId === "openai-codex" && isOpenAICodexCredentialRefreshRequired(e)) {
        await clearBrowserOAuthProvider(providerKeys, settings, providerId);
        console.warn("[auth] OpenAI (ChatGPT Plus/Pro): stored OAuth grant is stale; cleared credentials, please login again");
      } else {
        console.warn(`[auth] ${providerId}: failed to restore (${getErrorMessage(e)})`);
      }
    }
  }
}
