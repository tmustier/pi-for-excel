import type { OAuthCredentials, OAuthLoginCallbacks } from "@earendil-works/pi-ai";

export interface BrowserOAuthProvider {
  id: string;
  name: string;
  login(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials>;
  refreshToken(credentials: OAuthCredentials): Promise<OAuthCredentials>;
  getApiKey(credentials: OAuthCredentials): string;
}
