/**
 * Centralized mapping of OAuth provider IDs to API provider names.
 *
 * IMPORTANT: openai-codex uses chatgpt.com/backend-api, NOT api.openai.com.
 * Mapping it to "openai" would route requests to the wrong endpoint.
 */

const PROVIDER_MAP: Record<string, string> = {
  "anthropic": "anthropic",
  "openai-codex": "openai-codex",
  "github-copilot": "github-copilot",
  "gemini-cli": "google-gemini-cli",
  "google-gemini-cli": "google-gemini-cli",
  "antigravity": "google-antigravity",
  "google-antigravity": "google-antigravity",
};

export function mapToApiProvider(oauthProviderId: string): string {
  return PROVIDER_MAP[oauthProviderId] || oauthProviderId;
}

export const BROWSER_OAUTH_PROVIDERS = [
  "anthropic",
  "openai-codex",
  "google-gemini-cli",
  "google-antigravity",
  "github-copilot",
];
