/**
 * Browser-safe OpenAI Codex (ChatGPT OAuth) provider.
 *
 * `@earendil-works/pi-ai` ships an OpenAI Codex OAuth provider that relies on a
 * local Node callback server, which cannot run inside Office webviews.
 *
 * This implementation keeps the same OAuth endpoints/client config, then accepts
 * either an auto-captured local callback URL from the proxy helper or a manual
 * paste fallback (redirect URL or code#state) so login works in-browser.
 */

import type {
  OAuthCredentials,
  OAuthLoginCallbacks,
} from "@earendil-works/pi-ai";

import type { BrowserOAuthProvider } from "./browser-oauth-types.js";
import { generatePKCE } from "./pkce.js";

const CLIENT_ID = "app_EMoamEEZ73f0CkXaXp7hrann";
const AUTHORIZE_URL = "https://auth.openai.com/oauth/authorize";
const TOKEN_URL = "https://auth.openai.com/oauth/token";
const REDIRECT_URI = "http://localhost:1455/auth/callback";
const SCOPE = "openid profile email offline_access api.connectors.read api.connectors.invoke";
const ORIGINATOR = "codex_cli_rs";
const CREDENTIAL_VERSION = "codex-cli-rs-connector-scopes-2026-04";
const STALE_CREDENTIAL_ERROR = "OpenAI login needs to be refreshed for current Codex scopes";
const JWT_CLAIM_PATH = "https://api.openai.com/auth";

type ParsedAuthorizationInput = { code?: string; state?: string };

type TokenPayload = {
  idToken?: string;
  accessToken: string;
  refreshToken: string;
  expiresInSeconds: number;
};

type VersionedOpenAICodexCredentials = OAuthCredentials & {
  codexOAuthVersion?: string;
  scopes?: string;
};

function isOpenaiCodexBrowserOauthPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function createState(): string {
  const bytes = new Uint8Array(32);
  crypto.getRandomValues(bytes);

  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }

  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function createParsedAuthorizationInput(
  code: string | null | undefined,
  state: string | null | undefined,
): ParsedAuthorizationInput {
  const parsed: ParsedAuthorizationInput = {};
  if (code !== null && code !== undefined) {
    parsed.code = code;
  }
  if (state !== null && state !== undefined) {
    parsed.state = state;
  }
  return parsed;
}

function parseAuthorizationInput(input: string): ParsedAuthorizationInput {
  const value = input.trim();
  if (!value) return {};

  try {
    const url = new URL(value);
    return createParsedAuthorizationInput(
      url.searchParams.get("code"),
      url.searchParams.get("state"),
    );
  } catch {
    // not a URL
  }

  if (value.includes("#")) {
    const parts = value.split("#", 2);
    return createParsedAuthorizationInput(parts[0], parts[1]);
  }

  if (value.includes("code=")) {
    const params = new URLSearchParams(value.startsWith("?") ? value.slice(1) : value);
    return createParsedAuthorizationInput(params.get("code"), params.get("state"));
  }

  return { code: value };
}

function parseTokenPayload(payload: DynamicValue): TokenPayload | null {
  if (!isOpenaiCodexBrowserOauthPayloadShape(payload)) {
    return null;
  }

  const idToken = payload.id_token;
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

  const tokenPayload: TokenPayload = {
    accessToken,
    refreshToken,
    expiresInSeconds: expiresIn,
  };
  if (typeof idToken === "string") {
    tokenPayload.idToken = idToken;
  }
  return tokenPayload;
}

async function exchangeAuthorizationCode(code: string, verifier: string): Promise<TokenPayload> {
  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      code,
      redirect_uri: REDIRECT_URI,
      client_id: CLIENT_ID,
      code_verifier: verifier,
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

function decodeJwtPayload(token: string): DynamicObject | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) {
      return null;
    }

    const payloadSegment = parts[1] ?? "";
    const decoded = decodeBase64Url(payloadSegment);
    const parsed: DynamicValue = JSON.parse(decoded);

    if (!isOpenaiCodexBrowserOauthPayloadShape(parsed)) {
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
  if (!isOpenaiCodexBrowserOauthPayloadShape(authClaim)) {
    return null;
  }

  const accountId = authClaim.chatgpt_account_id;
  if (typeof accountId !== "string" || accountId.length === 0) {
    return null;
  }

  return accountId;
}

export function isOpenAICodexCredentialRefreshRequired(error: DynamicValue): boolean {
  return error instanceof Error && error.message.includes(STALE_CREDENTIAL_ERROR);
}

const REQUIRED_SCOPES = SCOPE.split(" ");

/**
 * Check whether the access token JWT itself proves the required scopes were
 * granted (`scp` array or space-delimited `scope`/`scp` string claim).
 *
 * Credentials imported from pi's `auth.json` (via the dev-only `/__pi-auth`
 * endpoint) or restored from older storage never carry the
 * `codexOAuthVersion`/`scopes` marker fields — those are only written by this
 * module's own login/refresh flows. Without this check, every externally
 * sourced Codex credential would be rejected as stale even when the
 * underlying grant already includes the connector scopes.
 */
function accessTokenHasRequiredScopes(accessToken: DynamicValue): boolean {
  if (typeof accessToken !== "string") {
    return false;
  }

  const payload = decodeJwtPayload(accessToken);
  if (!payload) {
    return false;
  }

  const claim = payload.scp ?? payload.scope;
  let granted: string[];
  if (Array.isArray(claim)) {
    granted = claim.filter((scope): scope is string => typeof scope === "string");
  } else if (typeof claim === "string") {
    granted = claim.split(" ");
  } else {
    return false;
  }

  return REQUIRED_SCOPES.every((scope) => granted.includes(scope));
}

function isCurrentOpenAICodexCredential(credentials: OAuthCredentials): boolean {
  const versioned = credentials as VersionedOpenAICodexCredentials;
  if (versioned.codexOAuthVersion === CREDENTIAL_VERSION && versioned.scopes === SCOPE) {
    return true;
  }

  return accessTokenHasRequiredScopes(credentials.access);
}

function assertCurrentOpenAICodexCredential(credentials: OAuthCredentials): void {
  if (!isCurrentOpenAICodexCredential(credentials)) {
    throw new Error(`${STALE_CREDENTIAL_ERROR}. Please disconnect and log in again.`);
  }
}

function buildOpenAICodexCredentials(tokens: TokenPayload, accountId: string): OAuthCredentials {
  const credentials: VersionedOpenAICodexCredentials = {
    access: tokens.accessToken,
    refresh: tokens.refreshToken,
    expires: Date.now() + tokens.expiresInSeconds * 1000,
    accountId,
    codexOAuthVersion: CREDENTIAL_VERSION,
    scopes: SCOPE,
  };

  return credentials;
}

async function createAuthorizationFlow(): Promise<{ verifier: string; state: string; url: string }> {
  const { verifier, challenge } = await generatePKCE(64);
  const state = createState();

  const authUrl = new URL(AUTHORIZE_URL);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("client_id", CLIENT_ID);
  authUrl.searchParams.set("redirect_uri", REDIRECT_URI);
  authUrl.searchParams.set("scope", SCOPE);
  authUrl.searchParams.set("code_challenge", challenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("id_token_add_organizations", "true");
  authUrl.searchParams.set("codex_cli_simplified_flow", "true");
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("originator", ORIGINATOR);

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
      "After login, Pi for Excel will try to capture the localhost callback automatically if the local proxy is running. " +
      "If it does not continue, copy the full callback URL from the browser address bar and paste it back in Pi for Excel.",
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
    throw new Error("OpenAI login failed: access token is missing ChatGPT account ID");
  }

  return buildOpenAICodexCredentials(tokens, accountId);
}

export async function refreshOpenAICodexBrowserToken(
  credentials: OAuthCredentials,
): Promise<OAuthCredentials> {
  assertCurrentOpenAICodexCredential(credentials);

  const refreshToken = credentials.refresh;
  if (typeof refreshToken !== "string" || refreshToken.trim().length === 0) {
    throw new Error("OpenAI refresh failed: missing refresh token");
  }

  const tokens = await refreshAccessToken(refreshToken);
  const accountId = getAccountId(tokens.accessToken);
  if (!accountId) {
    throw new Error("OpenAI refresh failed: access token is missing ChatGPT account ID");
  }

  return buildOpenAICodexCredentials(tokens, accountId);
}

export const openaiCodexBrowserOAuthProvider: BrowserOAuthProvider = {
  id: "openai-codex",
  name: "OpenAI (ChatGPT Plus/Pro)",

  async login(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
    return loginOpenAICodexInBrowser(callbacks);
  },

  async refreshToken(credentials: OAuthCredentials): Promise<OAuthCredentials> {
    return refreshOpenAICodexBrowserToken(credentials);
  },

  getApiKey(credentials: OAuthCredentials): string {
    assertCurrentOpenAICodexCredential(credentials);
    return credentials.access;
  },
};
