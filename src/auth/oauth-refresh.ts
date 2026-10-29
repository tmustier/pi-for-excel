import type { OAuthCredentials } from "@earendil-works/pi-ai";

import type { SettingsAccess } from "../storage/local/settings-store.js";
import type { ProviderKeysStoreLike } from "../storage/local/provider-credentials-store.js";
import type { BrowserOAuthProvider } from "./browser-oauth-types.js";
import { clearOAuthCredentials, loadOAuthCredentials, saveOAuthCredentials } from "./oauth-storage.js";
import { isOpenAICodexCredentialRefreshRequired } from "./openai-codex-browser-oauth.js";
import { BROWSER_OAUTH_PROVIDERS, mapToApiProvider } from "./provider-map.js";
import { getErrorMessage } from "../utils/errors.js";

export const OAUTH_REFRESH_MARGIN_MS = 2 * 60_000;

export type OAuthRefreshLock = <T>(name: string, run: () => Promise<T>) => Promise<T>;

export type OAuthRefreshOutcome =
  | "not_oauth"
  | "fresh"
  | "refreshed"
  | "logged_out"
  | "failed"
  | "expired";

export interface OAuthRefresher {
  ensureFresh(apiProvider: string): Promise<OAuthRefreshOutcome>;
}

function defaultLock(): OAuthRefreshLock {
  const locks = typeof navigator !== "undefined" ? navigator.locks : undefined;
  if (!locks) {
    return (_name, run) => run();
  }
  return (name, run) => locks.request(name, run);
}

export function createOAuthRefresher(options: {
  settings: SettingsAccess;
  providerKeys: Pick<ProviderKeysStoreLike, "set" | "delete">;
  getOAuthProvider: (providerId: string) => BrowserOAuthProvider | undefined;
  lock?: OAuthRefreshLock;
  now?: () => number;
  warn?: (message: string) => void;
}): OAuthRefresher {
  const lock = options.lock ?? defaultLock();
  const now = options.now ?? (() => Date.now());
  const warn = options.warn ?? ((message: string) => console.warn(message));
  const inFlight = new Map<string, Promise<OAuthRefreshOutcome>>();

  const needsRefresh = (credentials: OAuthCredentials): boolean =>
    now() >= credentials.expires - OAUTH_REFRESH_MARGIN_MS;

  const refreshUnderLock = async (oauthProviderId: string, apiProvider: string): Promise<OAuthRefreshOutcome> => {
    const provider = options.getOAuthProvider(oauthProviderId);
    if (!provider) return "not_oauth";

    return lock(`pi-oauth-refresh:${oauthProviderId}`, async () => {
      const credentials = await loadOAuthCredentials(options.settings, oauthProviderId);
      if (!credentials) return "not_oauth";
      if (!needsRefresh(credentials)) {
        await options.providerKeys.set(apiProvider, provider.getApiKey(credentials));
        return "fresh";
      }

      let refreshed: OAuthCredentials;
      try {
        refreshed = await provider.refreshToken(credentials);
      } catch (error) {
        if (oauthProviderId === "openai-codex" && isOpenAICodexCredentialRefreshRequired(error)) {
          await clearOAuthCredentials(options.settings, oauthProviderId).catch(() => {});
          await options.providerKeys.delete(apiProvider).catch(() => {});
          warn(`[auth] ${provider.name}: stored OAuth grant is stale; cleared credentials, please login again`);
          return "logged_out";
        }
        warn(`[auth] ${provider.name}: refresh failed (${getErrorMessage(error)})`);
        return now() < credentials.expires ? "failed" : "expired";
      }

      await saveOAuthCredentials(options.settings, oauthProviderId, refreshed);
      await options.providerKeys.set(apiProvider, provider.getApiKey(refreshed));
      return "refreshed";
    });
  };

  return {
    async ensureFresh(apiProvider: string): Promise<OAuthRefreshOutcome> {
      const oauthProviderId = BROWSER_OAUTH_PROVIDERS.find((id) => mapToApiProvider(id) === apiProvider);
      if (!oauthProviderId) return "not_oauth";

      const stored = await loadOAuthCredentials(options.settings, oauthProviderId);
      if (!stored) return "not_oauth";
      if (!needsRefresh(stored)) return "fresh";

      const pending = inFlight.get(oauthProviderId);
      if (pending) return pending;

      const run = refreshUnderLock(oauthProviderId, apiProvider).finally(() => {
        inFlight.delete(oauthProviderId);
      });
      inFlight.set(oauthProviderId, run);
      return run;
    },
  };
}
