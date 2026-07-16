import type { OAuthCredentials, OAuthLoginCallbacks } from "@earendil-works/pi-ai";

/**
 * Browser-owned OAuth adapter used by the existing provider-login UI.
 *
 * Pi AI 0.80.8 moved providers to the new AuthInteraction contract and no
 * longer exports the legacy coding-agent OAuthProviderInterface. Keeping this
 * small host interface lets the Office-safe OAuth flows remain independent of
 * Node callback-server implementations.
 */
export interface BrowserOAuthProvider {
  id: string;
  name: string;
  login(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials>;
  refreshToken(credentials: OAuthCredentials): Promise<OAuthCredentials>;
  getApiKey(credentials: OAuthCredentials): string;
}
