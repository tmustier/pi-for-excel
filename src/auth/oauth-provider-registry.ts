/**
 * Minimal OAuth provider registry for the Excel taskpane.
 *
 * We intentionally avoid importing `@mariozechner/pi-ai`'s OAuth index, which
 * registers CLI-only providers (Google Antigravity / Gemini CLI) that pull in
 * Node-only modules (e.g. `http`) and bloat browser bundles.
 */

import type { OAuthProviderInterface } from "@mariozechner/pi-ai";

import { anthropicOAuthProvider } from "@mariozechner/pi-ai/dist/utils/oauth/anthropic.js";
import { githubCopilotOAuthProvider } from "@mariozechner/pi-ai/dist/utils/oauth/github-copilot.js";

import { openaiCodexBrowserOAuthProvider } from "./openai-codex-browser-oauth.js";

const OAUTH_PROVIDERS: Map<string, OAuthProviderInterface> = new Map([
  [anthropicOAuthProvider.id, anthropicOAuthProvider],
  [openaiCodexBrowserOAuthProvider.id, openaiCodexBrowserOAuthProvider],
  [githubCopilotOAuthProvider.id, githubCopilotOAuthProvider],
]);

export function getOAuthProvider(id: string): OAuthProviderInterface | undefined {
  return OAUTH_PROVIDERS.get(id);
}
