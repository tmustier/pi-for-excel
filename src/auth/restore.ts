/**
 * Auto-restore auth credentials from pi's auth.json (dev) or localStorage (browser OAuth).
 *
 * Priority:
 * 1. pi's ~/.pi/agent/auth.json (served by Vite plugin at /__pi-auth)
 * 2. localStorage (credentials from in-browser OAuth flows)
 */

import { originalFetch } from "./cors-proxy.js";
import { mapToApiProvider, BROWSER_OAUTH_PROVIDERS } from "./provider-map.js";
import { getErrorMessage } from "../utils/errors.js";

/**
 * Restore credentials from all available sources.
 * Populates the ProviderKeysStore so ChatPanel can make API calls.
 */
export async function restoreCredentials(providerKeys: any): Promise<void> {
  const { getOAuthProvider } = await import("@mariozechner/pi-ai");

  // 1. Try pi's auth.json (dev server only)
  if (await restoreFromPiAuth(providerKeys, getOAuthProvider)) {
    return;
  }

  // 2. Fallback: localStorage (browser OAuth sessions)
  await restoreFromLocalStorage(providerKeys, getOAuthProvider);
}

async function restoreFromPiAuth(
  providerKeys: any,
  getOAuthProvider: any,
): Promise<boolean> {
  try {
    const res = await originalFetch("/__pi-auth");
    if (!res.ok) return false;

    const authData = await res.json() as Record<string, any>;
    console.log(`[auth] Found pi auth.json with ${Object.keys(authData).length} provider(s)`);

    for (const [providerId, cred] of Object.entries(authData)) {
      try {
        const apiProvider = mapToApiProvider(providerId);

        if (cred.type === "api_key" && cred.key) {
          await providerKeys.set(apiProvider, cred.key);
          console.log(`[auth] ${providerId}: API key loaded`);

        } else if (cred.type === "oauth") {
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

async function restoreFromLocalStorage(
  providerKeys: any,
  getOAuthProvider: any,
): Promise<void> {
  for (const providerId of BROWSER_OAUTH_PROVIDERS) {
    const stored = localStorage.getItem(`oauth_${providerId}`);
    if (!stored) continue;

    try {
      const credentials = JSON.parse(stored);
      const provider = getOAuthProvider(providerId);
      if (!provider) continue;

      const apiProvider = mapToApiProvider(providerId);

      if (Date.now() >= credentials.expires) {
        try {
          const refreshed = await provider.refreshToken(credentials);
          localStorage.setItem(`oauth_${providerId}`, JSON.stringify(refreshed));
          await providerKeys.set(apiProvider, provider.getApiKey(refreshed));
          console.log(`[auth] ${provider.name}: token refreshed from localStorage`);
        } catch (e: unknown) {
          console.warn(`[auth] ${provider.name}: refresh failed (${getErrorMessage(e)}), please login again`);
        }
      } else {
        await providerKeys.set(apiProvider, provider.getApiKey(credentials));
        console.log(`[auth] ${provider.name}: session restored from localStorage`);
      }
    } catch (e: unknown) {
      console.warn(`[auth] ${providerId}: failed to restore (${getErrorMessage(e)})`);
    }
  }
}
