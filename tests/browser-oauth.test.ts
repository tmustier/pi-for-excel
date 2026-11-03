import { authDeferred } from "./fixtures/auth-deferred.ts";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { test } from "node:test";
import { isCorsError } from "../src/auth/cors-error.ts";

import type { OAuthLoginCallbacks } from "@earendil-works/pi-ai";

import { getOAuthProvider } from "../src/auth/oauth-provider-registry.ts";
import { sha256ForPkce } from "../src/auth/pkce.ts";

function encodeBase64Url(value: string): string {
  return Buffer.from(value, "utf8")
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function fakeJwt(payload: Record<string, unknown>): string {
  return [
    encodeBase64Url(JSON.stringify({ alg: "none" })),
    encodeBase64Url(JSON.stringify(payload)),
    "signature",
  ].join(".");
}

function requestUrlToString(url: string | URL | Request): string {
  if (typeof url === "string") return url;
  if (url instanceof URL) return url.toString();
  return url.url;
}

function parseRequestBody(raw: unknown): Record<string, unknown> {
  if (typeof raw !== "object" || raw === null || Array.isArray(raw)) {
    throw new Error("request body must be an object");
  }
  return raw as Record<string, unknown>;
}

void test("PKCE SHA-256 falls back when Web Crypto digest is unavailable", async (t) => {
  const originalCrypto = globalThis.crypto;
  Object.defineProperty(globalThis, "crypto", {
    configurable: true,
    value: { getRandomValues: originalCrypto.getRandomValues.bind(originalCrypto) },
  });

  t.after(() => {
    Object.defineProperty(globalThis, "crypto", {
      configurable: true,
      value: originalCrypto,
    });
  });

  const input = new TextEncoder().encode("wps-http-pkce");
  const actual = Buffer.from(await sha256ForPkce(input)).toString("hex");
  const expected = createHash("sha256").update(input).digest("hex");

  assert.equal(actual, expected);
});

void test("Anthropic OAuth provider uses the browser-safe implementation", async (t) => {
  const originalFetch = globalThis.fetch;
  const requests: Array<{ url: string; body: unknown }> = [];

  globalThis.fetch = ((url: string | URL | Request, init?: RequestInit): Promise<Response> => {
    requests.push({ url: requestUrlToString(url), body: init?.body });
    return Promise.resolve(new Response(JSON.stringify({
      access_token: "sk-ant-oat-browser-test",
      refresh_token: "refresh-anthropic",
      expires_in: 3600,
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  });

  t.after(() => {
    globalThis.fetch = originalFetch;
  });

  const provider = getOAuthProvider("anthropic");
  assert.ok(provider);

  let authUrl = "";
  const callbacks: OAuthLoginCallbacks = {
    onAuth: (info) => {
      authUrl = info.url;
    },
    onPrompt: () => Promise.resolve(`anthropic-code#${new URL(authUrl).searchParams.get("state") ?? ""}`),
  };

  const credentials = await provider.login(callbacks);
  assert.equal(credentials.access, "sk-ant-oat-browser-test");
  assert.equal(credentials.refresh, "refresh-anthropic");

  const authorizeUrl = new URL(authUrl);
  assert.equal(authorizeUrl.hostname, "claude.ai");
  assert.equal(authorizeUrl.searchParams.get("client_id"), "9d1c250a-e61b-44d9-88ed-5944d1962f5e");
  assert.equal(authorizeUrl.searchParams.get("redirect_uri"), "http://localhost:53692/callback");

  assert.equal(requests[0]?.url, "https://platform.claude.com/v1/oauth/token");
  const body = parseRequestBody(JSON.parse(String(requests[0]?.body)) as unknown);
  assert.equal(body.grant_type, "authorization_code");
  assert.equal(body.code, "anthropic-code");
  assert.equal(body.state, new URL(authUrl).searchParams.get("state"));
});

void test("OpenAI Codex browser OAuth matches official Codex CLI authorize parameters", async (t) => {
  const originalFetch = globalThis.fetch;
  const requests: Array<{ url: string; body: unknown }> = [];
  const accessToken = fakeJwt({
    "https://api.openai.com/auth": {
      chatgpt_account_id: "acct_browser_test",
    },
  });

  globalThis.fetch = ((url: string | URL | Request, init?: RequestInit): Promise<Response> => {
    requests.push({ url: requestUrlToString(url), body: init?.body });
    return Promise.resolve(new Response(JSON.stringify({
      id_token: accessToken,
      access_token: accessToken,
      refresh_token: "refresh-openai",
      expires_in: 3600,
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  });

  t.after(() => {
    globalThis.fetch = originalFetch;
  });

  const provider = getOAuthProvider("openai-codex");
  assert.ok(provider);

  let authUrl = "";
  const callbacks: OAuthLoginCallbacks = {
    onAuth: (info) => {
      authUrl = info.url;
    },
    onPrompt: () => Promise.resolve(`openai-code#${new URL(authUrl).searchParams.get("state") ?? ""}`),
  };

  const credentials = await provider.login(callbacks);
  assert.equal(credentials.accountId, "acct_browser_test");
  assert.equal(credentials.codexOAuthVersion, "codex-cli-rs-connector-scopes-2026-04");
  assert.equal(
    credentials.scopes,
    "openid profile email offline_access api.connectors.read api.connectors.invoke",
  );

  const authorizeUrl = new URL(authUrl);
  assert.equal(authorizeUrl.hostname, "auth.openai.com");
  assert.equal(
    authorizeUrl.searchParams.get("scope"),
    "openid profile email offline_access api.connectors.read api.connectors.invoke",
  );
  assert.equal(authorizeUrl.searchParams.get("originator"), "codex_cli_rs");
  assert.equal(authorizeUrl.searchParams.get("codex_cli_simplified_flow"), "true");
  assert.equal(authorizeUrl.searchParams.get("id_token_add_organizations"), "true");
  assert.match(authorizeUrl.searchParams.get("state") ?? "", /^[A-Za-z0-9_-]{43}$/);

  assert.equal(requests[0]?.url, "https://auth.openai.com/oauth/token");
  const body = new URLSearchParams(String(requests[0]?.body));
  assert.equal(body.get("grant_type"), "authorization_code");
  assert.equal(body.get("code"), "openai-code");
  assert.equal(body.get("redirect_uri"), "http://localhost:1455/auth/callback");
  assert.equal(body.get("client_id"), "app_EMoamEEZ73f0CkXaXp7hrann");
  assert.match(body.get("code_verifier") ?? "", /^[A-Za-z0-9_-]{86}$/);
});

void test("OpenAI Codex browser OAuth accepts pi-sourced credentials with connector scopes in the JWT", async (t) => {
  const provider = getOAuthProvider("openai-codex");
  assert.ok(provider);

  const scopedAccessToken = fakeJwt({
    scp: ["openid", "profile", "email", "offline_access", "api.connectors.read", "api.connectors.invoke"],
    "https://api.openai.com/auth": {
      chatgpt_account_id: "acct_pi_auth_json",
    },
  });
  const piSourcedCredentials = {
    access: scopedAccessToken,
    refresh: "refresh-from-pi-auth-json",
    expires: Date.now() + 3600_000,
  };

  assert.equal(provider.getApiKey(piSourcedCredentials), scopedAccessToken);

  const originalFetch = globalThis.fetch;
  const requests: Array<{ url: string; body: unknown }> = [];
  globalThis.fetch = ((url: string | URL | Request, init?: RequestInit): Promise<Response> => {
    requests.push({ url: requestUrlToString(url), body: init?.body });
    return Promise.resolve(new Response(JSON.stringify({
      access_token: scopedAccessToken,
      refresh_token: "refresh-rotated",
      expires_in: 3600,
    }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));
  });
  t.after(() => {
    globalThis.fetch = originalFetch;
  });

  const refreshed = await provider.refreshToken(piSourcedCredentials);
  assert.equal(refreshed.refresh, "refresh-rotated");
  assert.equal(refreshed.accountId, "acct_pi_auth_json");
  assert.equal(
    (refreshed as { codexOAuthVersion?: string }).codexOAuthVersion,
    "codex-cli-rs-connector-scopes-2026-04",
  );

  const body = new URLSearchParams(String(requests[0]?.body));
  assert.equal(body.get("grant_type"), "refresh_token");
  assert.equal(body.get("refresh_token"), "refresh-from-pi-auth-json");
});

void test("OpenAI Codex browser OAuth accepts space-delimited scope claims", () => {
  const provider = getOAuthProvider("openai-codex");
  assert.ok(provider);

  const scopedAccessToken = fakeJwt({
    scope: "openid profile email offline_access api.connectors.read api.connectors.invoke",
    "https://api.openai.com/auth": {
      chatgpt_account_id: "acct_scope_string",
    },
  });

  assert.equal(
    provider.getApiKey({
      access: scopedAccessToken,
      refresh: "refresh-scope-string",
      expires: Date.now() + 3600_000,
    }),
    scopedAccessToken,
  );
});

void test("OpenAI Codex browser OAuth rejects JWT credentials missing connector scopes", async () => {
  const provider = getOAuthProvider("openai-codex");
  assert.ok(provider);

  const basicAccessToken = fakeJwt({
    scp: ["openid", "profile", "email", "offline_access"],
    "https://api.openai.com/auth": {
      chatgpt_account_id: "acct_basic_scopes",
    },
  });
  const staleCredentials = {
    access: basicAccessToken,
    refresh: "refresh-basic-scopes",
    expires: Date.now() + 3600_000,
  };

  assert.throws(
    () => provider.getApiKey(staleCredentials),
    /OpenAI login needs to be refreshed for current Codex scopes/,
  );

  await assert.rejects(
    () => provider.refreshToken(staleCredentials),
    /OpenAI login needs to be refreshed for current Codex scopes/,
  );
});

void test("OpenAI Codex browser OAuth rejects stale pre-scope-upgrade credentials", async () => {
  const provider = getOAuthProvider("openai-codex");
  assert.ok(provider);

  const staleCredentials = {
    access: "old-access-token",
    refresh: "old-refresh-token",
    expires: Date.now() + 3600_000,
  };

  assert.throws(
    () => provider.getApiKey(staleCredentials),
    /OpenAI login needs to be refreshed for current Codex scopes/,
  );

  await assert.rejects(
    () => provider.refreshToken(staleCredentials),
    /OpenAI login needs to be refreshed for current Codex scopes/,
  );
});

void test("OpenAI Codex browser OAuth requires account ID on the access token", async (t) => {
  const originalFetch = globalThis.fetch;
  const accessTokenWithoutAccount = fakeJwt({ sub: "user_without_embedded_account" });
  const idTokenWithAccount = fakeJwt({
    "https://api.openai.com/auth": {
      chatgpt_account_id: "acct_id_token_only",
    },
  });

  globalThis.fetch = (() => Promise.resolve(new Response(JSON.stringify({
    id_token: idTokenWithAccount,
    access_token: accessTokenWithoutAccount,
    refresh_token: "refresh-openai",
    expires_in: 3600,
  }), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  })));

  t.after(() => {
    globalThis.fetch = originalFetch;
  });

  const provider = getOAuthProvider("openai-codex");
  assert.ok(provider);

  let authUrl = "";
  const callbacks: OAuthLoginCallbacks = {
    onAuth: (info) => {
      authUrl = info.url;
    },
    onPrompt: () => Promise.resolve(`openai-code#${new URL(authUrl).searchParams.get("state") ?? ""}`),
  };

  await assert.rejects(
    () => provider.login(callbacks),
    /access token is missing ChatGPT account ID/,
  );
});

void test("a login flow whose code cannot load is not reported as a CORS problem", () => {
  const moduleFailure = new TypeError(
    "Failed to fetch dynamically imported module: https://localhost:3141/node_modules/.vite/deps/github-copilot.js",
  );
  assert.equal(isCorsError(moduleFailure), false);
  assert.equal(isCorsError(new TypeError("Failed to fetch")), true);
});

for (const providerId of ["anthropic", "openai-codex", "google-gemini-cli", "google-antigravity"]) {
  for (const cancelAt of ["prompt", "token", "before-exchange"] as const) {
    void test(`${providerId} cancellation at ${cancelAt} never returns credentials`, async (t) => {
      const provider = getOAuthProvider(providerId);
      assert.ok(provider);
      const abort = new AbortController();
      const originalFetch = globalThis.fetch;
      let requests = 0;
      globalThis.fetch = () => {
        requests += 1;
        abort.abort();
        return Promise.resolve(Response.json({
          access_token: fakeJwt({ "https://api.openai.com/auth": { chatgpt_account_id: "fixture-account" } }),
          refresh_token: "fixture-refresh",
          expires_in: 3600,
        }));
      };
      t.after(() => { globalThis.fetch = originalFetch; });
      await assert.rejects(provider.login({
        signal: abort.signal,
        onAuth: () => {},
        onPrompt: () => {
          if (cancelAt === "prompt") abort.abort();
          return Promise.resolve("fixture-code");
        },
        onProgress: (message) => {
          if (cancelAt === "before-exchange" && /Exchanging/.test(message)) abort.abort();
        },
      }), { name: "AbortError" });
      assert.equal(requests, cancelAt === "token" ? 1 : 0);
    });
  }
}

void test("Google cancellation during account fetch never starts project discovery", async (t) => {
  const abort = new AbortController();
  const provider = getOAuthProvider("google-gemini-cli");
  assert.ok(provider);
  const originalFetch = globalThis.fetch;
  const requests: string[] = [];
  globalThis.fetch = (input) => {
    const url = requestUrlToString(input);
    requests.push(url);
    if (url.includes("userinfo")) abort.abort();
    return Promise.resolve(Response.json(url.includes("userinfo") ? { email: "fixture@example.test" } : {
      access_token: "fixture-access", refresh_token: "fixture-refresh", expires_in: 3600,
    }));
  };
  t.after(() => { globalThis.fetch = originalFetch; });
  await assert.rejects(provider.login({
    signal: abort.signal, onAuth: () => {}, onPrompt: () => Promise.resolve("fixture-code"),
  }), { name: "AbortError" });
  assert.equal(requests.length, 2);
});

for (const cancel of [false, true]) {
  void test(`Google project provisioning ${cancel ? "cancellation" : "deadline"} aborts a pending poll`, async (t) => {
    t.mock.timers.enable({ apis: ["setTimeout"] });
    const abort = new AbortController();
    const provider = getOAuthProvider("google-gemini-cli");
    assert.ok(provider);
    const originalFetch = globalThis.fetch;
    const polling = authDeferred<void>();
    let pollAborted = false;
    globalThis.fetch = (input, init) => {
      const url = requestUrlToString(input);
      if (url.endsWith("/operations/fixture")) {
        const signal = init?.signal;
        signal?.addEventListener("abort", () => { pollAborted = true; }, { once: true });
        polling.resolve();
        return new Promise((_resolve, reject) => {
          signal?.addEventListener("abort", () => reject(signal.reason instanceof Error ? signal.reason : new Error("Aborted")), { once: true });
        });
      }
      return Promise.resolve(Response.json(url.includes("/token")
        ? { access_token: "fixture-access", refresh_token: "fixture-refresh", expires_in: 3600 }
        : url.includes("loadCodeAssist")
          ? { allowedTiers: [{ id: "free-tier", isDefault: true }] }
          : url.includes("onboardUser") ? { name: "operations/fixture", done: false } : {}));
    };
    t.after(() => { globalThis.fetch = originalFetch; });
    const login = provider.login({
      signal: abort.signal, onAuth: () => {}, onPrompt: () => Promise.resolve("fixture-code"),
    });
    const rejected = assert.rejects(login, cancel ? { name: "AbortError" } : /provisioning timed out/);
    await polling.promise;
    if (cancel) abort.abort();
    else t.mock.timers.tick(2 * 60_000);
    assert.equal(pollAborted, true, "project provisioning must abort the in-flight poll");
    await rejected;
  });
}
