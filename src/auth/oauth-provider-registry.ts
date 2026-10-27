/**
 * Minimal OAuth provider registry for the Excel taskpane.
 *
 * Pi AI 0.80.8 replaced the legacy OAuthProviderInterface with provider-owned
 * AuthInteraction flows. The taskpane still curates browser-safe providers, so
 * this module adapts the upstream GitHub flow while keeping our Office-safe
 * Google, OpenAI and Anthropic implementations.
 */

import type {
  OAuthCredential,
  OAuthCredentials,
  OAuthLoginCallbacks,
} from "@earendil-works/pi-ai";
// pi-ai loads its OAuth flows through an opaque import() so bundlers skip
// Node-only flows; Vite cannot follow it, and the taskpane failed with
// "Failed to fetch dynamically imported module" (LOGIN-21, P-65). The GitHub
// Copilot flow itself only uses fetch, so it is imported directly. It is not in
// pi-ai's export map, hence the file path; check it on pi-ai upgrades.
import { githubCopilotOAuth } from "../../node_modules/@earendil-works/pi-ai/dist/auth/oauth/github-copilot.js";

import { anthropicBrowserOAuthProvider } from "./anthropic-browser-oauth.js";
import type { BrowserOAuthProvider } from "./browser-oauth-types.js";
import { openaiCodexBrowserOAuthProvider } from "./openai-codex-browser-oauth.js";
import {
  googleAntigravityBrowserOAuthProvider,
  googleGeminiCliBrowserOAuthProvider,
} from "./google-browser-oauth.js";

function fromOAuthCredential(credential: OAuthCredential): OAuthCredentials {
  const copied: OAuthCredentials = { ...credential };
  delete copied.type;
  return copied;
}

const githubCopilotBrowserOAuthProvider: BrowserOAuthProvider = {
  id: "github-copilot",
  name: githubCopilotOAuth.name,
  async login(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
    return fromOAuthCredential(await githubCopilotOAuth.login({
      signal: callbacks.signal ?? new AbortController().signal,
      prompt: async (prompt) => {
        if (prompt.type === "select") {
          const selection = await callbacks.onSelect({
            message: prompt.message,
            options: prompt.options.map((option) => ({ id: option.id, label: option.label })),
          });
          if (typeof selection !== "string") throw new Error("Login cancelled");
          return selection;
        }
        if (prompt.type === "manual_code" && callbacks.onManualCodeInput) {
          return callbacks.onManualCodeInput();
        }
        return callbacks.onPrompt({
          message: prompt.message,
          ...(prompt.placeholder !== undefined ? { placeholder: prompt.placeholder } : {}),
          allowEmpty: prompt.type === "text",
        });
      },
      notify: (event) => {
        if (event.type === "auth_url") {
          callbacks.onAuth({
            url: event.url,
            ...(event.instructions !== undefined ? { instructions: event.instructions } : {}),
          });
          return;
        }
        if (event.type === "device_code") {
          callbacks.onDeviceCode({
            userCode: event.userCode,
            verificationUri: event.verificationUri,
            ...(event.intervalSeconds !== undefined ? { intervalSeconds: event.intervalSeconds } : {}),
            ...(event.expiresInSeconds !== undefined ? { expiresInSeconds: event.expiresInSeconds } : {}),
          });
          return;
        }
        callbacks.onProgress?.(event.message);
      },
    }));
  },
  async refreshToken(credentials: OAuthCredentials): Promise<OAuthCredentials> {
    return fromOAuthCredential(
      await githubCopilotOAuth.refresh({ ...credentials, type: "oauth" }, new AbortController().signal),
    );
  },
  getApiKey(credentials: OAuthCredentials): string {
    return credentials.access;
  },
};

const OAUTH_PROVIDERS: Map<string, BrowserOAuthProvider> = new Map([
  [anthropicBrowserOAuthProvider.id, anthropicBrowserOAuthProvider],
  [openaiCodexBrowserOAuthProvider.id, openaiCodexBrowserOAuthProvider],
  [googleGeminiCliBrowserOAuthProvider.id, googleGeminiCliBrowserOAuthProvider],
  [googleAntigravityBrowserOAuthProvider.id, googleAntigravityBrowserOAuthProvider],
  [githubCopilotBrowserOAuthProvider.id, githubCopilotBrowserOAuthProvider],
]);

export function getOAuthProvider(id: string): BrowserOAuthProvider | undefined {
  return OAUTH_PROVIDERS.get(id);
}
