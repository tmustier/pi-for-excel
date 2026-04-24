/**
 * Browser-safe Anthropic (Claude Pro/Max) OAuth provider.
 *
 * pi-ai's built-in Anthropic OAuth provider starts a Node callback server, which
 * is not available inside the Office taskpane WebView. This mirrors the same
 * authorization-code + PKCE flow but uses the existing manual paste prompt.
 */

import type {
  OAuthCredentials,
  OAuthLoginCallbacks,
  OAuthProviderInterface,
} from "@mariozechner/pi-ai";

import { generatePKCE } from "./pkce.js";

const CLIENT_ID = "9d1c250a-e61b-44d5-88ed-5944d1962f5e";
const AUTHORIZE_URL = "https://claude.ai/oauth/authorize";
const TOKEN_URL = "https://platform.claude.com/v1/oauth/token";
const REDIRECT_URI = "http://localhost:53692/callback";
const SCOPES = "org:create_api_key user:profile user:inference user:sessions:claude_code user:mcp_servers user:file_upload";

type ParsedAuthorizationInput = { code?: string; state?: string };

type TokenPayload = {
  accessToken: string;
  refreshToken: string;
  expiresInSeconds: number;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseAuthorizationInput(input: string): ParsedAuthorizationInput {
  const value = input.trim();
  if (!value) return {};

  try {
    const url = new URL(value);
    return {
      code: url.searchParams.get("code") ?? undefined,
      state: url.searchParams.get("state") ?? undefined,
    };
  } catch {
    // not a URL
  }

  if (value.includes("#")) {
    const [code, state] = value.split("#", 2);
    return { code, state };
  }

  if (value.includes("code=")) {
    const params = new URLSearchParams(value.startsWith("?") ? value.slice(1) : value);
    return {
      code: params.get("code") ?? undefined,
      state: params.get("state") ?? undefined,
    };
  }

  return { code: value };
}

function parseTokenPayload(payload: unknown): TokenPayload | null {
  if (!isRecord(payload)) return null;

  const accessToken = payload.access_token;
  const refreshToken = payload.refresh_token;
  const expiresIn = payload.expires_in;

  if (
    typeof accessToken !== "string" ||
    typeof refreshToken !== "string" ||
    typeof expiresIn !== "number"
  ) {
    return null;
  }

  return {
    accessToken,
    refreshToken,
    expiresInSeconds: expiresIn,
  };
}

async function postTokenRequest(body: Record<string, string>, failurePrefix: string): Promise<TokenPayload> {
  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    throw new Error(`${failurePrefix} (${response.status}): ${errorText}`);
  }

  const payload = parseTokenPayload(await response.json());
  if (!payload) {
    throw new Error(`${failurePrefix}: response is missing required fields`);
  }

  return payload;
}

async function exchangeAuthorizationCode(code: string, state: string, verifier: string): Promise<TokenPayload> {
  return postTokenRequest(
    {
      grant_type: "authorization_code",
      client_id: CLIENT_ID,
      code,
      state,
      redirect_uri: REDIRECT_URI,
      code_verifier: verifier,
    },
    "Anthropic token exchange failed",
  );
}

async function refreshAccessToken(refreshToken: string): Promise<TokenPayload> {
  return postTokenRequest(
    {
      grant_type: "refresh_token",
      client_id: CLIENT_ID,
      refresh_token: refreshToken,
    },
    "Anthropic token refresh failed",
  );
}

async function createAuthorizationFlow(): Promise<{ verifier: string; url: string }> {
  const { verifier, challenge } = await generatePKCE();
  const authUrl = new URL(AUTHORIZE_URL);
  authUrl.searchParams.set("code", "true");
  authUrl.searchParams.set("client_id", CLIENT_ID);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", REDIRECT_URI);
  authUrl.searchParams.set("scope", SCOPES);
  authUrl.searchParams.set("code_challenge", challenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("state", verifier);

  return { verifier, url: authUrl.toString() };
}

export async function loginAnthropicInBrowser(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
  callbacks.onProgress?.("Preparing Anthropic login…");
  const flow = await createAuthorizationFlow();

  callbacks.onAuth({
    url: flow.url,
    instructions:
      "Complete login in your browser. If the browser cannot reach localhost, copy the final redirect URL and paste it back in Pi for Excel.",
  });

  if (callbacks.signal?.aborted) {
    throw new Error("Login cancelled");
  }

  callbacks.onProgress?.("Waiting for authorization code…");
  const input = await callbacks.onPrompt({
    message: "Paste the authorization code or full redirect URL:",
    placeholder: REDIRECT_URI,
  });

  const parsed = parseAuthorizationInput(input);
  if (parsed.state && parsed.state !== flow.verifier) {
    throw new Error("Anthropic login failed: OAuth state mismatch");
  }
  if (!parsed.code) {
    throw new Error("Anthropic login failed: missing authorization code");
  }

  callbacks.onProgress?.("Exchanging code for tokens…");
  const tokens = await exchangeAuthorizationCode(parsed.code, parsed.state ?? flow.verifier, flow.verifier);

  return {
    access: tokens.accessToken,
    refresh: tokens.refreshToken,
    expires: Date.now() + tokens.expiresInSeconds * 1000 - 5 * 60 * 1000,
  };
}

export async function refreshAnthropicBrowserToken(credentials: OAuthCredentials): Promise<OAuthCredentials> {
  const refreshToken = credentials.refresh;
  if (typeof refreshToken !== "string" || refreshToken.trim().length === 0) {
    throw new Error("Anthropic refresh failed: missing refresh token");
  }

  const tokens = await refreshAccessToken(refreshToken);
  return {
    access: tokens.accessToken,
    refresh: tokens.refreshToken,
    expires: Date.now() + tokens.expiresInSeconds * 1000 - 5 * 60 * 1000,
  };
}

export const anthropicBrowserOAuthProvider: OAuthProviderInterface = {
  id: "anthropic",
  name: "Anthropic (Claude Pro/Max)",

  async login(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
    return loginAnthropicInBrowser(callbacks);
  },

  async refreshToken(credentials: OAuthCredentials): Promise<OAuthCredentials> {
    return refreshAnthropicBrowserToken(credentials);
  },

  getApiKey(credentials: OAuthCredentials): string {
    return credentials.access;
  },
};
