/**
 * On-demand refresh of browser OAuth logins.
 *
 * A request asks for a fresh access token just before it is sent. A token that
 * expires within the margin is refreshed once: concurrent requests in this
 * taskpane share one refresh, and other taskpanes on the same machine wait on
 * a Web Lock and then reuse the token the first one stored. Refreshing rotates
 * the refresh token, so two taskpanes refreshing the same grant at once would
 * leave one of them holding a revoked token.
 */

import type { OAuthCredentials } from "@earendil-works/pi-ai";

import type { SettingsAccess } from "../storage/local/settings-store.js";
import type { ProviderKeysStoreLike } from "../storage/local/provider-credentials-store.js";
import type { BrowserOAuthProvider } from "./browser-oauth-types.js";
import { clearOAuthCredentials, loadOAuthCredentials, saveOAuthCredentials } from "./oauth-storage.js";
import { isOpenAICodexCredentialRefreshRequired } from "./openai-codex-browser-oauth.js";
import { BROWSER_OAUTH_PROVIDERS, mapToApiProvider } from "./provider-map.js";
import { getErrorMessage } from "../utils/errors.js";

/** Refresh this long before expiry so a request never leaves with a token that dies in flight. */
export const OAUTH_REFRESH_MARGIN_MS = 2 * 60_000;

/** Cross-taskpane mutual exclusion. Web Locks are shared by every taskpane of one browser profile. */
export type OAuthRefreshLock = <T>(name: string, run: () => Promise<T>) => Promise<T>;

export type OAuthRefreshOutcome =
  /** Not a browser OAuth login, or no stored grant: nothing to do. */
  | "not_oauth"
  /** The stored token is still valid (possibly refreshed by another taskpane). */
  | "fresh"
  | "refreshed"
  /** The grant was revoked or is stale; the login was removed and must be redone. */
  | "logged_out"
  /** Refresh failed for another reason (e.g. offline); the old token is kept. */
  | "failed";

export interface OAuthRefresher {
  /** Makes sure the API key stored for `apiProvider` is usable for the next request. */
  ensureFresh(apiProvider: string): Promise<OAuthRefreshOutcome>;
}

function defaultLock(): OAuthRefreshLock {
  const locks = typeof navigator !== "undefined" ? navigator.locks : undefined;
  if (!locks) {
    // Without Web Locks only this taskpane is serialized (single flight below).
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
      // Re-read inside the lock: another taskpane may have refreshed meanwhile.
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
        return "failed";
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

      // Cheap path: no lock when the stored token is comfortably valid.
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
