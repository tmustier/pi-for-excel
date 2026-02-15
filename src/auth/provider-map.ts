/**
 * Centralized mapping of OAuth provider IDs to API provider names.
 *
 * IMPORTANT: openai-codex uses chatgpt.com/backend-api, NOT api.openai.com.
 * Mapping it to "openai" would route requests to the wrong endpoint.
 */

const PROVIDER_MAP: Record<string, string> = {
  "anthropic": "anthropic",
  "openai-codex": "openai-codex",   // chatgpt.com/backend-api, NOT api.openai.com
  "github-copilot": "github-copilot",
  "gemini-cli": "google",
  "google-gemini-cli": "google",
  "antigravity": "google",
  "google-antigravity": "google",
};

/** Map an OAuth provider ID to the API provider name used by pi-ai */
export function mapToApiProvider(oauthProviderId: string): string {
  return PROVIDER_MAP[oauthProviderId] || oauthProviderId;
}

/** OAuth providers whose flows work in-browser (PKCE/manual paste, no local callback server) */
export const BROWSER_OAUTH_PROVIDERS = ["anthropic", "openai-codex", "github-copilot"];
