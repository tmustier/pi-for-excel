/**
 * Shared browser-safe Google OAuth flow for Office taskpanes.
 */

import type {
  OAuthCredentials,
  OAuthLoginCallbacks,
} from "@earendil-works/pi-ai";

import type { BrowserOAuthProvider } from "./browser-oauth-types.js";
import { generatePKCE } from "./pkce.js";

const GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";

type ParsedAuthorizationInput = { code?: string; state?: string };

type GoogleTokenPayload = {
  accessToken: string;
  refreshToken: string;
  expiresInSeconds: number;
};

export type GoogleOAuthFlowConfig = {
  id: string;
  name: string;
  clientId: string;
  clientSecret: string;
  redirectUri: string;
  scopes: readonly string[];
  discoverProject: (
    accessToken: string,
    callbacks: OAuthLoginCallbacks,
  ) => Promise<string>;
};

function isGoogleBrowserOauthCorePayloadShape(value: DynamicValue): value is DynamicObject {
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

function parseTokenPayload(
  payload: DynamicValue,
  fallbackRefreshToken?: string,
): GoogleTokenPayload | null {
  if (!isGoogleBrowserOauthCorePayloadShape(payload)) {
    return null;
  }

  const accessToken = payload.access_token;
  const refreshTokenValue = payload.refresh_token;
  const expiresIn = payload.expires_in;

  const refreshToken =
    typeof refreshTokenValue === "string"
      ? refreshTokenValue
      : fallbackRefreshToken;

  if (
    typeof accessToken !== "string"
    || typeof refreshToken !== "string"
    || refreshToken.trim().length === 0
    || typeof expiresIn !== "number"
  ) {
    return null;
  }

  return {
    accessToken,
    refreshToken,
    expiresInSeconds: expiresIn,
  };
}

async function exchangeAuthorizationCode(
  config: GoogleOAuthFlowConfig,
  code: string,
  verifier: string,
): Promise<GoogleTokenPayload> {
  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      code,
      grant_type: "authorization_code",
      redirect_uri: config.redirectUri,
      code_verifier: verifier,
    }),
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    throw new Error(`Google token exchange failed (${response.status}): ${errorText}`);
  }

  const payload = parseTokenPayload(await response.json());
  if (!payload) {
    throw new Error("Google token response is missing required fields");
  }

  return payload;
}

async function refreshAccessToken(
  config: GoogleOAuthFlowConfig,
  refreshToken: string,
): Promise<GoogleTokenPayload> {
  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      refresh_token: refreshToken,
      grant_type: "refresh_token",
    }),
  });

  if (!response.ok) {
    const errorText = await response.text().catch(() => "");
    throw new Error(`Google token refresh failed (${response.status}): ${errorText}`);
  }

  const payload = parseTokenPayload(await response.json(), refreshToken);
  if (!payload) {
    throw new Error("Google refresh response is missing required fields");
  }

  return payload;
}

async function getUserEmail(accessToken: string): Promise<string | undefined> {
  try {
    const response = await fetch("https://www.googleapis.com/oauth2/v1/userinfo?alt=json", {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!response.ok) {
      return undefined;
    }

    const payload: DynamicValue = await response.json();
    if (!isGoogleBrowserOauthCorePayloadShape(payload)) {
      return undefined;
    }

    const email = payload.email;
    if (typeof email !== "string" || email.trim().length === 0) {
      return undefined;
    }

    return email;
  } catch {
    return undefined;
  }
}

async function loginGoogleOAuth(
  config: GoogleOAuthFlowConfig,
  callbacks: OAuthLoginCallbacks,
): Promise<OAuthCredentials> {
  callbacks.onProgress?.("Preparing Google login…");

  const { verifier, challenge } = await generatePKCE();
  const state = createState();

  const authUrl = new URL(GOOGLE_AUTH_URL);
  authUrl.searchParams.set("client_id", config.clientId);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("redirect_uri", config.redirectUri);
  authUrl.searchParams.set("scope", config.scopes.join(" "));
  authUrl.searchParams.set("code_challenge", challenge);
  authUrl.searchParams.set("code_challenge_method", "S256");
  authUrl.searchParams.set("state", state);
  authUrl.searchParams.set("access_type", "offline");
  authUrl.searchParams.set("prompt", "consent");

  callbacks.onAuth({
    url: authUrl.toString(),
    instructions:
      "After sign-in, Pi for Excel will try to capture the localhost callback automatically if the local proxy is running. " +
      "If it does not continue, copy the full callback URL from the browser address bar and paste it back in Pi for Excel.",
  });

  if (callbacks.signal?.aborted) {
    throw new Error("Login cancelled");
  }

  callbacks.onProgress?.("Waiting for authorization code…");
  const input = await callbacks.onPrompt({
    message: "Paste the authorization code (or full redirect URL):",
    placeholder: `${config.redirectUri}?code=...&state=...`,
  });

  const parsed = parseAuthorizationInput(input);
  if (parsed.state && parsed.state !== state) {
    throw new Error("Google login failed: OAuth state mismatch");
  }

  if (!parsed.code) {
    throw new Error("Google login failed: missing authorization code");
  }

  callbacks.onProgress?.("Exchanging authorization code for tokens…");
  const tokens = await exchangeAuthorizationCode(config, parsed.code, verifier);

  callbacks.onProgress?.("Fetching account details…");
  const email = await getUserEmail(tokens.accessToken);

  const projectId = await config.discoverProject(tokens.accessToken, callbacks);

  const credentials: OAuthCredentials = {
    refresh: tokens.refreshToken,
    access: tokens.accessToken,
    expires: Date.now() + tokens.expiresInSeconds * 1000 - 5 * 60 * 1000,
    projectId,
  };
  if (email !== undefined) {
    credentials.email = email;
  }
  return credentials;
}

async function refreshGoogleOAuth(
  config: GoogleOAuthFlowConfig,
  credentials: OAuthCredentials,
): Promise<OAuthCredentials> {
  const refreshToken = credentials.refresh;
  if (typeof refreshToken !== "string" || refreshToken.trim().length === 0) {
    throw new Error("Google refresh failed: missing refresh token");
  }

  const projectId = credentials.projectId;
  if (typeof projectId !== "string" || projectId.trim().length === 0) {
    throw new Error("Google refresh failed: missing project ID");
  }

  const tokens = await refreshAccessToken(config, refreshToken);

  const refreshed: OAuthCredentials = {
    refresh: tokens.refreshToken,
    access: tokens.accessToken,
    expires: Date.now() + tokens.expiresInSeconds * 1000 - 5 * 60 * 1000,
    projectId,
  };
  if (credentials.email !== undefined) {
    refreshed.email = credentials.email;
  }
  return refreshed;
}

export function createGoogleBrowserOAuthProvider(
  config: GoogleOAuthFlowConfig,
): BrowserOAuthProvider {
  return {
    id: config.id,
    name: config.name,

    async login(callbacks: OAuthLoginCallbacks): Promise<OAuthCredentials> {
      return loginGoogleOAuth(config, callbacks);
    },

    async refreshToken(credentials: OAuthCredentials): Promise<OAuthCredentials> {
      return refreshGoogleOAuth(config, credentials);
    },

    getApiKey(credentials: OAuthCredentials): string {
      const projectId = credentials.projectId;
      if (typeof projectId !== "string" || projectId.trim().length === 0) {
        throw new Error("Google OAuth credentials are missing projectId");
      }

      return JSON.stringify({
        token: credentials.access,
        projectId,
      });
    },
  };
}
