import type { OAuthLoginCallbacks } from "@earendil-works/pi-ai";

import {
  type GoogleOAuthFlowConfig,
  createGoogleBrowserOAuthProvider,
} from "./google-browser-oauth-core.js";

const decode = (value: string): string => atob(value);

const GEMINI_CLI_CLIENT_ID = decode(
  "NjgxMjU1ODA5Mzk1LW9vOGZ0Mm9wcmRybnA5ZTNh"
  + "cWY2YXYzaG1kaWIxMzVqLmFwcHMuZ29vZ2xldXNlcmNvbnRlbnQuY29t",
);
const GEMINI_CLI_CLIENT_SECRET = decode(
  "R09DU1BYLTR1SGdNUG0tMW83"
  + "U2stZ2VWNkN1NWNsWEZzeGw=",
);
const GEMINI_CLI_REDIRECT_URI = "http://localhost:8085/oauth2callback";

const ANTIGRAVITY_CLIENT_ID = decode(
  "MTA3MTAwNjA2MDU5MS10bWhzc2luMmgyMWxjcmUyMzV2"
  + "dG9sb2poNGc0MDNlcC5hcHBzLmdvb2dsZXVzZXJjb250ZW50LmNvbQ==",
);
const ANTIGRAVITY_CLIENT_SECRET = decode(
  "R09DU1BYLUs1OEZXUjQ4"
  + "NkxkTEoxbUxCOHNYQzR6NnFEQWY=",
);
const ANTIGRAVITY_REDIRECT_URI = "http://localhost:51121/oauth-callback";

const GEMINI_CLI_SCOPES = [
  "https://www.googleapis.com/auth/cloud-platform",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/userinfo.profile",
];

const ANTIGRAVITY_SCOPES = [
  "https://www.googleapis.com/auth/cloud-platform",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/userinfo.profile",
  "https://www.googleapis.com/auth/cclog",
  "https://www.googleapis.com/auth/experimentsandconfigs",
];

const CODE_ASSIST_ENDPOINT = "https://cloudcode-pa.googleapis.com";
const ANTIGRAVITY_ENDPOINTS = [
  "https://cloudcode-pa.googleapis.com",
  "https://daily-cloudcode-pa.sandbox.googleapis.com",
] as const;
const ANTIGRAVITY_DEFAULT_PROJECT_ID = "rising-fact-p41fc";

const TIER_FREE = "free-tier";
const TIER_LEGACY = "legacy-tier";
const TIER_STANDARD = "standard-tier";

type CodeAssistHeaders = {
  Authorization: string;
  "Content-Type": string;
  "User-Agent": string;
  "X-Goog-Api-Client": string;
  "Client-Metadata"?: string;
};

function isAuthGoogleBrowserOauthPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function getProjectId(data: unknown): string | undefined {
  if (!isAuthGoogleBrowserOauthPayloadShape(data)) return undefined;

  const direct = data.cloudaicompanionProject;
  if (typeof direct === "string" && direct.trim().length > 0) {
    return direct;
  }

  if (isAuthGoogleBrowserOauthPayloadShape(direct)) {
    const nestedId = direct.id;
    if (typeof nestedId === "string" && nestedId.trim().length > 0) {
      return nestedId;
    }
  }

  return undefined;
}

function getDefaultTierId(data: unknown): string {
  if (!isAuthGoogleBrowserOauthPayloadShape(data)) return TIER_LEGACY;

  const allowedTiers = data.allowedTiers;
  if (!Array.isArray(allowedTiers)) return TIER_LEGACY;

  let fallbackTier: string | undefined;
  for (const tier of allowedTiers) {
    if (!isAuthGoogleBrowserOauthPayloadShape(tier)) continue;

    const tierId = tier.id;
    const isDefault = tier.isDefault;

    if (!fallbackTier && typeof tierId === "string" && tierId.trim().length > 0) {
      fallbackTier = tierId;
    }

    if (isDefault === true && typeof tierId === "string" && tierId.trim().length > 0) {
      return tierId;
    }
  }

  return fallbackTier ?? TIER_LEGACY;
}

function hasCurrentTier(data: unknown): boolean {
  if (!isAuthGoogleBrowserOauthPayloadShape(data)) return false;
  return isAuthGoogleBrowserOauthPayloadShape(data.currentTier);
}

function isVpcScAffectedUser(payload: unknown): boolean {
  if (!isAuthGoogleBrowserOauthPayloadShape(payload)) return false;

  const error = payload.error;
  if (!isAuthGoogleBrowserOauthPayloadShape(error)) return false;

  const details = error.details;
  if (!Array.isArray(details)) return false;

  for (const detail of details) {
    if (!isAuthGoogleBrowserOauthPayloadShape(detail)) continue;
    if (detail.reason === "SECURITY_POLICY_VIOLATED") {
      return true;
    }
  }

  return false;
}

function wait(ms: number, signal: AbortSignal): Promise<void> {
  signal.throwIfAborted();
  return new Promise((resolve, reject) => {
    const onAbort = (): void => {
      clearTimeout(timer);
      reject(signal.reason instanceof Error ? signal.reason : new Error("Google project provisioning aborted"));
    };
    const timer = setTimeout(() => {
      signal.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    signal.addEventListener("abort", onAbort, { once: true });
  });
}

function buildCodeAssistHeaders(accessToken: string): CodeAssistHeaders {
  return {
    Authorization: `Bearer ${accessToken}`,
    "Content-Type": "application/json",
    "User-Agent": "google-api-nodejs-client/9.15.1",
    "X-Goog-Api-Client": "google-cloud-sdk vscode_cloudshelleditor/0.1",
    "Client-Metadata": JSON.stringify({
      ideType: "IDE_UNSPECIFIED",
      platform: "PLATFORM_UNSPECIFIED",
      pluginType: "GEMINI",
    }),
  };
}

async function promptForProjectId(callbacks: OAuthLoginCallbacks): Promise<string> {
  const raw = await callbacks.onPrompt({
    message:
      "Enter your Google Cloud project ID (required for this account/workspace tier):",
    placeholder: "my-google-cloud-project-id",
  });

  callbacks.signal?.throwIfAborted();
  const projectId = raw.trim();
  if (!projectId) {
    throw new Error("Google login failed: project ID is required");
  }

  return projectId;
}

async function pollOnboardingOperation(
  operationName: string,
  headers: CodeAssistHeaders,
  callbacks: OAuthLoginCallbacks,
): Promise<unknown> {
  let attempt = 0;
  const timeout = new AbortController();
  const timer = setTimeout(() => timeout.abort(new Error("Google project provisioning timed out")), 2 * 60_000);
  const signal = callbacks.signal ? AbortSignal.any([callbacks.signal, timeout.signal]) : timeout.signal;

  try {
    while (true) {
      signal.throwIfAborted();
      if (attempt > 0) {
        callbacks.onProgress?.(`Waiting for project provisioning (attempt ${attempt + 1})…`);
        await wait(5000, signal);
      }

      const response = await fetch(`${CODE_ASSIST_ENDPOINT}/v1internal/${operationName}`, {
        method: "GET",
        signal,
        headers,
      });

      if (!response.ok) {
        throw new Error(`Failed to poll Google onboarding operation: ${response.status} ${response.statusText}`);
      }

      const payload: unknown = await response.json();
      signal.throwIfAborted();
      if (!isAuthGoogleBrowserOauthPayloadShape(payload)) {
        throw new Error("Google onboarding returned an invalid response payload");
      }

      if (payload.done === true) {
        return payload;
      }

      attempt += 1;
    }
  } finally {
    clearTimeout(timer);
  }
}

async function discoverGeminiCliProject(
  accessToken: string,
  callbacks: OAuthLoginCallbacks,
): Promise<string> {
  const headers = buildCodeAssistHeaders(accessToken);

  callbacks.onProgress?.("Checking for existing Cloud Code Assist project…");

  const loadResponse = await fetch(`${CODE_ASSIST_ENDPOINT}/v1internal:loadCodeAssist`, {
    method: "POST",
    signal: callbacks.signal ?? null,
    headers,
    body: JSON.stringify({
      metadata: {
        ideType: "IDE_UNSPECIFIED",
        platform: "PLATFORM_UNSPECIFIED",
        pluginType: "GEMINI",
      },
    }),
  });

  let loadData: unknown;
  if (!loadResponse.ok) {
    let errorPayload: unknown;
    try {
      errorPayload = await loadResponse.clone().json();
    } catch {
      errorPayload = undefined;
    }

    if (isVpcScAffectedUser(errorPayload)) {
      loadData = { currentTier: { id: TIER_STANDARD } };
    } else {
      const errorText = await loadResponse.text().catch(() => "");
      throw new Error(
        `Google loadCodeAssist failed (${loadResponse.status} ${loadResponse.statusText}): ${errorText}`,
      );
    }
  } else {
    loadData = await loadResponse.json();
  }

  callbacks.signal?.throwIfAborted();
  const existingProjectId = getProjectId(loadData);
  if (existingProjectId) {
    return existingProjectId;
  }

  if (hasCurrentTier(loadData)) {
    return promptForProjectId(callbacks);
  }

  const tierId = getDefaultTierId(loadData);
  const requiresProjectId = tierId !== TIER_FREE;
  const projectId = requiresProjectId ? await promptForProjectId(callbacks) : undefined;

  callbacks.signal?.throwIfAborted();
  callbacks.onProgress?.("Provisioning Cloud Code Assist project…");

  const onboardBody = {
    tierId,
    metadata: {
      ideType: "IDE_UNSPECIFIED",
      platform: "PLATFORM_UNSPECIFIED",
      pluginType: "GEMINI",
      ...(projectId ? { duetProject: projectId } : {}),
    },
    ...(projectId ? { cloudaicompanionProject: projectId } : {}),
  };

  const onboardResponse = await fetch(`${CODE_ASSIST_ENDPOINT}/v1internal:onboardUser`, {
    method: "POST",
    signal: callbacks.signal ?? null,
    headers,
    body: JSON.stringify(onboardBody),
  });

  if (!onboardResponse.ok) {
    const errorText = await onboardResponse.text().catch(() => "");
    throw new Error(
      `Google onboardUser failed (${onboardResponse.status} ${onboardResponse.statusText}): ${errorText}`,
    );
  }

  let onboardingData: unknown = await onboardResponse.json();
  callbacks.signal?.throwIfAborted();
  if (isAuthGoogleBrowserOauthPayloadShape(onboardingData) && onboardingData.done !== true) {
    const operationName = onboardingData.name;
    if (typeof operationName === "string" && operationName.trim().length > 0) {
      onboardingData = await pollOnboardingOperation(operationName, headers, callbacks);
    }
  }

  const onboardedProjectId =
    isAuthGoogleBrowserOauthPayloadShape(onboardingData) && isAuthGoogleBrowserOauthPayloadShape(onboardingData.response)
      ? getProjectId(onboardingData.response)
      : undefined;

  if (onboardedProjectId) {
    return onboardedProjectId;
  }

  if (projectId) {
    return projectId;
  }

  throw new Error("Google login failed: could not determine Cloud Code Assist project");
}

async function discoverAntigravityProject(
  accessToken: string,
  callbacks: OAuthLoginCallbacks,
): Promise<string> {
  const headers = buildCodeAssistHeaders(accessToken);

  callbacks.onProgress?.("Checking for existing Antigravity project…");

  for (const endpoint of ANTIGRAVITY_ENDPOINTS) {
    callbacks.signal?.throwIfAborted();
    try {
      const response = await fetch(`${endpoint}/v1internal:loadCodeAssist`, {
        method: "POST",
        signal: callbacks.signal ?? null,
        headers,
        body: JSON.stringify({
          metadata: {
            ideType: "IDE_UNSPECIFIED",
            platform: "PLATFORM_UNSPECIFIED",
            pluginType: "GEMINI",
          },
        }),
      });

      if (!response.ok) {
        continue;
      }

      const payload: unknown = await response.json();
      callbacks.signal?.throwIfAborted();
      const projectId = getProjectId(payload);
      if (projectId) {
        return projectId;
      }
    } catch {
      callbacks.signal?.throwIfAborted();
    }
  }

  callbacks.onProgress?.("Using default Antigravity project");
  return ANTIGRAVITY_DEFAULT_PROJECT_ID;
}

const GEMINI_CLI_CONFIG: GoogleOAuthFlowConfig = {
  id: "google-gemini-cli",
  name: "Google Cloud Code Assist (Gemini)",
  clientId: GEMINI_CLI_CLIENT_ID,
  clientSecret: GEMINI_CLI_CLIENT_SECRET,
  redirectUri: GEMINI_CLI_REDIRECT_URI,
  scopes: GEMINI_CLI_SCOPES,
  discoverProject: discoverGeminiCliProject,
};

const ANTIGRAVITY_CONFIG: GoogleOAuthFlowConfig = {
  id: "google-antigravity",
  name: "Google Antigravity",
  clientId: ANTIGRAVITY_CLIENT_ID,
  clientSecret: ANTIGRAVITY_CLIENT_SECRET,
  redirectUri: ANTIGRAVITY_REDIRECT_URI,
  scopes: ANTIGRAVITY_SCOPES,
  discoverProject: discoverAntigravityProject,
};

export const googleGeminiCliBrowserOAuthProvider = createGoogleBrowserOAuthProvider(GEMINI_CLI_CONFIG);
export const googleAntigravityBrowserOAuthProvider = createGoogleBrowserOAuthProvider(ANTIGRAVITY_CONFIG);
