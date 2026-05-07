/**
 * Stub for `@earendil-works/pi-ai/dist/utils/oauth/index.js` in browser builds.
 *
 * `pi-ai`'s OAuth index registers several CLI-only providers (e.g. Google
 * Antigravity / Gemini CLI) and sets up Node.js fetch proxy plumbing.
 *
 * The Excel taskpane uses its own minimal OAuth implementation (Anthropic +
 * GitHub Copilot) and should not bundle Node-only OAuth flows.
 */

export function getOAuthProvider(_id: string): undefined {
  return undefined;
}

export function registerOAuthProvider(_provider: unknown): void {
  // no-op
}

export function getOAuthProviders(): unknown[] {
  return [];
}

/** @deprecated */
export function getOAuthProviderInfoList(): Array<{ id: string; name: string; available: boolean }> {
  return [];
}

/** @deprecated */
export function refreshOAuthToken(_providerId: string, _credentials: unknown): Promise<never> {
  return Promise.reject(new Error("OAuth token refresh is not available in this build"));
}

export function getOAuthApiKey(
  _providerId: string,
  _credentials: Record<string, unknown>,
): Promise<null> {
  return Promise.resolve(null);
}
