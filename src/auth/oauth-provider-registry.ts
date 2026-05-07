/**
 * Minimal OAuth provider registry for the Excel taskpane.
 *
 * pi-ai 0.56+ exposes OAuth providers via the public `@earendil-works/pi-ai/oauth`
 * subpath instead of package-internal deep imports. We still curate our own
 * registry here so the taskpane can substitute browser-safe Google/OpenAI flows.
 */

import type { OAuthProviderInterface } from "@earendil-works/pi-ai";

import {
  githubCopilotOAuthProvider,
} from "@earendil-works/pi-ai/oauth";

import { anthropicBrowserOAuthProvider } from "./anthropic-browser-oauth.js";
import { openaiCodexBrowserOAuthProvider } from "./openai-codex-browser-oauth.js";
import {
  googleAntigravityBrowserOAuthProvider,
  googleGeminiCliBrowserOAuthProvider,
} from "./google-browser-oauth.js";

const OAUTH_PROVIDERS: Map<string, OAuthProviderInterface> = new Map([
  [anthropicBrowserOAuthProvider.id, anthropicBrowserOAuthProvider],
  [openaiCodexBrowserOAuthProvider.id, openaiCodexBrowserOAuthProvider],
  [googleGeminiCliBrowserOAuthProvider.id, googleGeminiCliBrowserOAuthProvider],
  [googleAntigravityBrowserOAuthProvider.id, googleAntigravityBrowserOAuthProvider],
  [githubCopilotOAuthProvider.id, githubCopilotOAuthProvider],
]);

export function getOAuthProvider(id: string): OAuthProviderInterface | undefined {
  return OAUTH_PROVIDERS.get(id);
}
