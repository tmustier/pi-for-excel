/**
 * Minimal OAuth provider registry for the Excel taskpane.
 *
 * Pi AI 0.80.8 replaced the legacy OAuthProviderInterface with provider-owned
 * AuthInteraction flows. The taskpane still curates browser-safe providers, so
 * this module adapts the upstream GitHub flow while keeping our Office-safe
 * Google, OpenAI and Anthropic implementations.
 */

import type {
  AuthEvent,
  AuthInteraction,
  AuthPrompt,
  OAuthCredential,
  OAuthCredentials,
  OAuthLoginCallbacks,
  OAuthSelectOption,
} from "@earendil-works/pi-ai";
import { githubCopilotProvider } from "@earendil-works/pi-ai/providers/github-copilot";

import { anthropicBrowserOAuthProvider } from "./anthropic-browser-oauth.js";
import type { BrowserOAuthProvider } from "./browser-oauth-types.js";
import { openaiCodexBrowserOAuthProvider } from "./openai-codex-browser-oauth.js";
import {
  googleAntigravityBrowserOAuthProvider,
  googleGeminiCliBrowserOAuthProvider,
} from "./google-browser-oauth.js";

function toOAuthCredential(credentials: OAuthCredentials): OAuthCredential {
  return {
    ...credentials,
    type: "oauth",
  };
}

function fromOAuthCredential(credential: OAuthCredential): OAuthCredentials {
  const copied: OAuthCredentials = { ...credential };
  delete copied.type;
  return copied;
}

function toLegacySelectOptions(
  options: AuthPrompt & { type: "select" },
): OAuthSelectOption[] {
  return options.options.map((option) => ({
    id: option.id,
    label: option.label,
  }));
}

async function handleAuthPrompt(
  callbacks: OAuthLoginCallbacks,
  prompt: AuthPrompt,
): Promise<string> {
  if (prompt.type === "select") {
    const selection = await callbacks.onSelect({
      message: prompt.message,
      options: toLegacySelectOptions(prompt),
    });
    if (typeof selection !== "string") {
      throw new Error("Login cancelled");
    }
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
}

function handleAuthEvent(callbacks: OAuthLoginCallbacks, event: AuthEvent): void {
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
}

function createAuthInteraction(callbacks: OAuthLoginCallbacks): AuthInteraction {
  return {
    ...(callbacks.signal !== undefined ? { signal: callbacks.signal } : {}),
    prompt: (prompt) => handleAuthPrompt(callbacks, prompt),
    notify: (event) => handleAuthEvent(callbacks, event),
  };
}

function createGitHubCopilotBrowserProvider(): BrowserOAuthProvider {
  const oauth = githubCopilotProvider().auth.oauth;
  if (!oauth) {
    throw new Error("GitHub Copilot OAuth is unavailable.");
  }

  return {
    id: "github-copilot",
    name: oauth.name,
    async login(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
      return fromOAuthCredential(await oauth.login(createAuthInteraction(callbacks)));
    },
    async refreshToken(credentials: OAuthCredentials): Promise<OAuthCredentials> {
      return fromOAuthCredential(await oauth.refresh(toOAuthCredential(credentials)));
    },
    getApiKey(credentials: OAuthCredentials): string {
      return credentials.access;
    },
  };
}

const githubCopilotBrowserOAuthProvider = createGitHubCopilotBrowserProvider();

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
