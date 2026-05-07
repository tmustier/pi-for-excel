import assert from "node:assert/strict";
import { test } from "node:test";

import type { OAuthLoginCallbacks } from "@earendil-works/pi-ai";

import { getOAuthProvider } from "../src/auth/oauth-provider-registry.ts";

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
  }) as typeof fetch;

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

  assert.equal(requests.length, 1);
  assert.equal(requests[0]?.url, "https://platform.claude.com/v1/oauth/token");
  const body = JSON.parse(String(requests[0]?.body)) as Record<string, unknown>;
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
  }) as typeof fetch;

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

  assert.equal(requests.length, 1);
  assert.equal(requests[0]?.url, "https://auth.openai.com/oauth/token");
  const body = new URLSearchParams(String(requests[0]?.body));
  assert.equal(body.get("grant_type"), "authorization_code");
  assert.equal(body.get("code"), "openai-code");
  assert.equal(body.get("redirect_uri"), "http://localhost:1455/auth/callback");
  assert.equal(body.get("client_id"), "app_EMoamEEZ73f0CkXaXp7hrann");
  assert.match(body.get("code_verifier") ?? "", /^[A-Za-z0-9_-]{86}$/);
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
  }))) as typeof fetch;

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
