/**
 * Browser-safe OpenAI Codex (ChatGPT OAuth) provider.
 *
 * `@mariozechner/pi-ai` ships an OpenAI Codex OAuth provider that relies on a
 * local Node callback server, which cannot run inside Office webviews.
 *
 * This implementation keeps the same OAuth endpoints/client config, but uses a
 * manual paste flow (redirect URL or code#state) so login works in-browser.
 */

import type {
  OAuthCredentials,
  OAuthLoginCallbacks,
  OAuthProviderInterface,
} from "@mariozechner/pi-ai";
import { generatePKCE } from "@mariozechner/pi-ai/dist/utils/oauth/pkce.js";

const CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann";
const AUTHORIZE_URL = "https://auth.openai.com/oauth/authorize";
const TOKEN_URL = "https://auth.openai.com/oauth/token";
const REDIRECT_URI = "http://localhost:1455/auth/callback";
const SCOPE = "openid profile email offline_access";
const JWT_CLAIM_PATH = "https://api.openai.com/auth";

type ParsedAuthorizationInput = { code?: string; state?: string };

type TokenPayload = {
  accessToken: string;
  refreshToken: string;
  expiresInSeconds: number;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function createState(): string {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);

  let out = "";
  for (const byte of bytes) {
    out += byte.toString(16).padStart(2, "0");
  }
  return out;
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
  if (!isRecord(payload)) {
    return null;
  }

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

async function exchangeAuthorizationCode(code: string, verifier: string): Promise<TokenPayload> {
  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      client_id: CLIENT_ID,
      code,
      code_verifier: verifier,
      redirect_uri: REDIRECT_URI,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    throw new Error(`OpenAI token exchange failed (${response.status}): ${errorText}`);
  }

  const payload = parseTokenPayload(await response.json());
  if (!payload) {
    throw new Error("OpenAI token response is missing required fields");
  }

  return payload;
}

async function refreshAccessToken(refreshToken: string): Promise<TokenPayload> {
  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      client_id: CLIENT_ID,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    throw new Error(`OpenAI token refresh failed (${response.status}): ${errorText}`);
  }

  const payload = parseTokenPayload(await response.json());
  if (!payload) {
    throw new Error("OpenAI refresh response is missing required fields");
  }

  return payload;
}

function decodeBase64Url(encoded: string): string {
  const normalized = encoded.replace(/-/g, "+").replace(/_/g, "/");
  const paddingLength = normalized.length % 4 === 0 ? 0 : 4 - (normalized.length % 4);
  const padded = normalized + "=".repeat(paddingLength);
  return atob(padded);
}

function decodeJwtPayload(token: string): Record<string, unknown> | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) {
      return null;
    }

    const payloadSegment = parts[1] ?? "";
    const decoded = decodeBase64Url(payloadSegment);
    const parsed: unknown = JSON.parse(decoded);

    if (!isRecord(parsed)) {
      return null;
    }

    return parsed;
  } catch {
    return null;
  }
}

function getAccountId(accessToken: string): string | null {
  const payload = decodeJwtPayload(accessToken);
  if (!payload) {
    return null;
  }

  const authClaim = payload[JWT_CLAIM_PATH];
  if (!isRecord(authClaim)) {
    return null;
  }

  const accountId = authClaim.chatgpt_account_id;
  if (typeof accountId !== "string" || accountId.length === 0) {
    return null;
  }

  return accountId;
}

async function createAuthorizationFlow(): Promise<{ verifier: string; state: string; url: string }> {
  const { verifier, challenge } = await generatePKCE();
  const state = createState();

  const authUrl = new URL(AUTHORIZE_URL);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("client_id", CLIENT_ID);
  authUrl.searchParams.set("redirect_uri", REDIRECT_URI);
  authUrl.searchParams.set("scope", SCOPE);
  authUrl.searchParams.set("code_challenge", challenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("id_token_add_organizations", "true");
  authUrl.searchParams.set("codex_cli_simplified_flow", "true");
  authUrl.searchParams.set("originator", "pi");

  return {
    verifier,
    state,
    url: authUrl.toString(),
  };
}

export async function loginOpenAICodexInBrowser(
  callbacks: OAuthLoginCallbacks,
): Promise<OAuthCredentials> {
  callbacks.onProgress?.("Preparing OpenAI login…");
  const flow = await createAuthorizationFlow();

  callbacks.onAuth({
    url: flow.url,
    instructions:
      "After login, copy the final URL from the browser address bar and paste it in Pi (it should contain code= and state=).",
  });

  if (callbacks.signal?.aborted) {
    throw new Error("Login cancelled");
  }

  callbacks.onProgress?.("Waiting for authorization code…");
  const input = await callbacks.onPrompt({
    message: "Paste the authorization code (or full redirect URL):",
    placeholder: "http://localhost:1455/auth/callback?code=...&state=...",
  });

  const parsed = parseAuthorizationInput(input);
  if (parsed.state && parsed.state !== flow.state) {
    throw new Error("OpenAI login failed: OAuth state mismatch");
  }

  if (!parsed.code) {
    throw new Error("OpenAI login failed: missing authorization code");
  }

  callbacks.onProgress?.("Exchanging code for tokens…");
  const tokens = await exchangeAuthorizationCode(parsed.code, flow.verifier);

  const accountId = getAccountId(tokens.accessToken);
  if (!accountId) {
    throw new Error("OpenAI login failed: could not extract ChatGPT account ID");
  }

  return {
    access: tokens.accessToken,
    refresh: tokens.refreshToken,
    expires: Date.now() + tokens.expiresInSeconds * 1000,
    accountId,
  };
}

export async function refreshOpenAICodexBrowserToken(
  credentials: OAuthCredentials,
): Promise<OAuthCredentials> {
  const refreshToken = credentials.refresh;
  if (typeof refreshToken !== "string" || refreshToken.trim().length === 0) {
    throw new Error("OpenAI refresh failed: missing refresh token");
  }

  const tokens = await refreshAccessToken(refreshToken);
  const accountId = getAccountId(tokens.accessToken);
  if (!accountId) {
    throw new Error("OpenAI refresh failed: could not extract ChatGPT account ID");
  }

  return {
    access: tokens.accessToken,
    refresh: tokens.refreshToken,
    expires: Date.now() + tokens.expiresInSeconds * 1000,
    accountId,
  };
}

export const openaiCodexBrowserOAuthProvider: OAuthProviderInterface = {
  id: "openai-codex",
  name: "OpenAI (ChatGPT Plus/Pro)",

  async login(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
    return loginOpenAICodexInBrowser(callbacks);
  },

  async refreshToken(credentials: OAuthCredentials): Promise<OAuthCredentials> {
    return refreshOpenAICodexBrowserToken(credentials);
  },

  getApiKey(credentials: OAuthCredentials): string {
    return credentials.access;
  },
};
