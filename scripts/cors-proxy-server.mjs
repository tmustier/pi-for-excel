#!/usr/bin/env node

/**
 * Minimal CORS proxy for Pi for Excel.
 *
 * Why this exists:
 * - Some provider OAuth/token endpoints (and some LLM APIs) block browser requests via CORS.
 * - In dev we rely on Vite's proxy. In production, you can run this locally and point
 *   Pi for Excel's proxy setting at it (default: https://localhost:3003; if
 *   3003 is busy and PORT is not set, the helper chooses a random free port).
 *
 * Usage:
 *   npm run proxy:https   # HTTPS (recommended for Office webviews)
 *   npm run proxy         # HTTP  (may be blocked as mixed content)
 *
 * Proxy format:
 *   https://<listen-host>:<listen-port>/?url=<target-url>
 *
 * Example:
 *   curl 'https://localhost:3003/?url=https%3A%2F%2Fexample.com'
 */

import http from "node:http";
import https from "node:https";
import fs from "node:fs";
import path from "node:path";
import { lookup as dnsLookup } from "node:dns/promises";
import { Readable } from "node:stream";

import {
  bridgeCodexWebSocketToSse,
  CODEX_WEBSOCKET_BRIDGE_HEADER,
  CODEX_WEBSOCKET_BRIDGE_TRANSPORT,
  isCodexWebSocketBridgeTarget,
} from "./codex-websocket-bridge.mjs";
import {
  evaluateTargetHostPolicy,
  isIpLiteral,
  normalizeHost,
  parseAllowedTargetHosts,
} from "./proxy-target-policy.mjs";
import {
  isAllowedClientAddress,
  parseClientCidrAllowlist,
} from "./proxy-client-policy.mjs";

const args = new Set(process.argv.slice(2));
const useHttps = args.has("--https") || process.env.HTTPS === "1" || process.env.HTTPS === "true";
const useHttp = args.has("--http");

if (useHttps && useHttp) {
  console.error("[pi-for-excel] Invalid args: can't use both --https and --http");
  process.exit(1);
}

const DEFAULT_PORT = 3003;
const hasExplicitHost = typeof process.env.HOST === "string" && process.env.HOST.trim().length > 0;
const HOST = hasExplicitHost ? process.env.HOST.trim() : (useHttps ? "localhost" : "127.0.0.1");
const LISTEN_HOSTS = hasExplicitHost
  ? [HOST]
  : useHttps
    ? ["127.0.0.1", "::1"]
    : [HOST];
const hasExplicitPort = typeof process.env.PORT === "string" && process.env.PORT.trim().length > 0;

function parsePort(rawPort) {
  const port = Number.parseInt(rawPort, 10);
  if (!Number.isInteger(port) || port < 0 || port > 65535) {
    console.error(`[pi-for-excel] Invalid PORT: ${rawPort}`);
    console.error("[pi-for-excel] Expected an integer from 0 to 65535.");
    process.exit(1);
  }
  return port;
}

const PORT = hasExplicitPort ? parsePort(process.env.PORT) : DEFAULT_PORT;

const OAUTH_CALLBACK_HOST = process.env.OAUTH_CALLBACK_HOST || "localhost";
const OAUTH_CALLBACK_SERVER_ENABLED = !["0", "false", "no"].includes(
  (process.env.OAUTH_CALLBACK_SERVER || "").trim().toLowerCase(),
);
const OAUTH_CALLBACK_CAPTURE_TTL_MS = 10 * 60 * 1000;
const MAX_OAUTH_CALLBACK_CAPTURES = 32;
const OAUTH_CALLBACK_PROVIDER_CONFIGS = [
  {
    providerId: "openai-codex",
    label: "OpenAI ChatGPT",
    path: "/auth/callback",
    port: parsePort(process.env.OPENAI_OAUTH_CALLBACK_PORT || "1455"),
  },
  {
    providerId: "anthropic",
    label: "Anthropic",
    path: "/callback",
    port: parsePort(process.env.ANTHROPIC_OAUTH_CALLBACK_PORT || "53692"),
  },
  {
    providerId: "google-gemini-cli",
    label: "Google Code Assist",
    path: "/oauth2callback",
    port: parsePort(process.env.GOOGLE_GEMINI_CLI_OAUTH_CALLBACK_PORT || "8085"),
  },
  {
    providerId: "google-antigravity",
    label: "Google Antigravity",
    path: "/oauth-callback",
    port: parsePort(process.env.GOOGLE_ANTIGRAVITY_OAUTH_CALLBACK_PORT || "51121"),
  },
];

const rootDir = path.resolve(process.cwd());
// Central deployments (docs/central-proxy.md) can point at org-issued certs.
const keyPath = process.env.TLS_KEY_PATH || path.join(rootDir, "key.pem");
const certPath = process.env.TLS_CERT_PATH || path.join(rootDir, "cert.pem");

// SECURITY: loopback-only by default. Central deployments must opt in to
// specific IPv4 client ranges; invalid entries are fatal (fail closed).
const allowedClientCidrs = (() => {
  const raw = process.env.ALLOWED_CLIENT_CIDRS;
  if (!raw || raw.trim().length === 0) return [];

  const { cidrs, invalid } = parseClientCidrAllowlist(raw);
  if (invalid.length > 0) {
    console.error(`[pi-for-excel] Invalid ALLOWED_CLIENT_CIDRS entries: ${invalid.join(", ")}`);
    console.error("[pi-for-excel] Expected comma-separated IPv4 CIDRs (e.g. 10.96.0.0/13) or bare IPv4 addresses. /0 is not allowed.");
    process.exit(1);
  }
  if (cidrs.length === 0) {
    console.error("[pi-for-excel] ALLOWED_CLIENT_CIDRS was set but contained no valid entries.");
    process.exit(1);
  }
  return cidrs;
})();

const HOP_BY_HOP_HEADERS = new Set([
  "connection",
  "keep-alive",
  "proxy-authenticate",
  "proxy-authorization",
  "te",
  "trailer",
  "transfer-encoding",
  "upgrade",
]);

// SECURITY: local CORS proxies are a common footgun. Even if bound to localhost,
// a browser tab on any origin can still call it unless we restrict CORS.
// Default allowlist matches our dev + hosted origins; override via env var.
const DEFAULT_ALLOWED_ORIGINS = new Set([
  "https://localhost:3141",
  "https://pi-for-excel.vercel.app",
]);

const allowedOrigins = (() => {
  const raw = process.env.ALLOWED_ORIGINS;
  if (!raw) return DEFAULT_ALLOWED_ORIGINS;
  const set = new Set(
    raw
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean),
  );
  return set.size > 0 ? set : DEFAULT_ALLOWED_ORIGINS;
})();

function isAllowedOrigin(origin) {
  return typeof origin === "string" && allowedOrigins.has(origin);
}

function envFlag(name) {
  const raw = process.env[name];
  return raw === "1" || raw === "true";
}

const DEFAULT_ALLOWED_TARGET_HOSTS = new Set([
  "api.anthropic.com",
  "console.anthropic.com",
  "platform.claude.com",
  "github.com",
  "api.github.com",
  "auth.openai.com",
  "api.openai.com",
  "chatgpt.com",
  "oauth2.googleapis.com",
  "generativelanguage.googleapis.com",
  "cloudcode-pa.googleapis.com",
  "daily-cloudcode-pa.sandbox.googleapis.com",
  "api.z.ai",
  // Web search providers
  "s.jina.ai",
  "api.firecrawl.dev",
  "google.serper.dev",
  "api.tavily.com",
  "api.search.brave.com",
]);

const allowAllTargetHosts = envFlag("ALLOW_ALL_TARGET_HOSTS");
const allowLoopbackTargets = envFlag("ALLOW_LOOPBACK_TARGETS");
const allowPrivateTargets = envFlag("ALLOW_PRIVATE_TARGETS");
const strictTargetResolution = envFlag("STRICT_TARGET_RESOLUTION");

const hasConfiguredAllowedTargetHosts =
  typeof process.env.ALLOWED_TARGET_HOSTS === "string"
  && process.env.ALLOWED_TARGET_HOSTS.trim().length > 0;

const configuredAllowedTargetHosts = hasConfiguredAllowedTargetHosts
  ? parseAllowedTargetHosts(process.env.ALLOWED_TARGET_HOSTS)
  : new Set();

// SECURITY: fail closed on explicit-but-unparseable ALLOWED_TARGET_HOSTS.
// Falling back to the default allowlist would silently re-enable legacy
// override semantics (loopback/private bypass, GitHub-enterprise path
// bypass) that a configured central proxy relies on being off.
if (hasConfiguredAllowedTargetHosts && configuredAllowedTargetHosts.size === 0) {
  console.error("[pi-for-excel] ALLOWED_TARGET_HOSTS was set but contained no valid host entries.");
  console.error("[pi-for-excel] Expected comma-separated hostnames or IP literals (e.g. api.deepseek.com,10.97.193.77).");
  process.exit(1);
}

const allowedTargetHosts = (() => {
  if (allowAllTargetHosts) {
    return new Set();
  }

  if (configuredAllowedTargetHosts.size > 0) {
    return configuredAllowedTargetHosts;
  }

  return new Set(DEFAULT_ALLOWED_TARGET_HOSTS);
})();

const EMPTY_ALLOWED_TARGET_HOSTS = new Set();

const TARGET_POLICY_MESSAGES = {
  blocked_target_invalid_host: "Invalid target host",
  blocked_target_not_allowlisted:
    "Target host is not allowlisted. Configure ALLOWED_TARGET_HOSTS or set ALLOW_ALL_TARGET_HOSTS=1 to disable host allowlisting.",
  blocked_target_loopback: "Loopback target URLs are blocked by default. Set ALLOW_LOOPBACK_TARGETS=1 to override.",
  blocked_target_private_ip: "Private/local target URLs are blocked by default. Set ALLOW_PRIVATE_TARGETS=1 to override.",
  blocked_target_resolution_failed: "Target hostname could not be resolved (STRICT_TARGET_RESOLUTION=1)",
};

function isGitHubEnterpriseOAuthPathname(pathname) {
  return pathname === "/login/device/code" || pathname === "/login/oauth/access_token";
}

function isGitHubEnterpriseCopilotPathname(pathname) {
  return pathname.startsWith("/copilot_internal/");
}

function shouldBypassHostAllowlistForGitHubEnterprise(targetUrl) {
  const hostname = normalizeHost(targetUrl.hostname);
  if (!hostname || isIpLiteral(hostname)) return false;

  if (isGitHubEnterpriseOAuthPathname(targetUrl.pathname)) {
    return hostname !== "github.com";
  }

  if (isGitHubEnterpriseCopilotPathname(targetUrl.pathname)) {
    if (hostname === "api.github.com" || hostname === "api.individual.githubcopilot.com") {
      return false;
    }

    return hostname.startsWith("api.");
  }

  return false;
}

function setCorsHeaders(req, res) {
  const origin = req.headers.origin;
  if (isAllowedOrigin(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
  }

  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,PATCH,DELETE,OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    req.headers["access-control-request-headers"] || "*",
  );
  res.setHeader(
    "Access-Control-Expose-Headers",
    `*, X-Pi-For-Excel-Proxy, ${CODEX_WEBSOCKET_BRIDGE_HEADER}`,
  );
  res.setHeader("Access-Control-Max-Age", "86400");
}

function rejectWithReason(res, reason) {
  const msg = TARGET_POLICY_MESSAGES[reason] || "forbidden";
  res.statusCode = 403;
  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.end(`${reason}: ${msg}`);
}

const oauthCallbackCaptures = new Map();

function findOAuthCallbackProvider(providerId) {
  return OAUTH_CALLBACK_PROVIDER_CONFIGS.find((config) => config.providerId === providerId);
}

function oauthCallbackKey(providerId, state) {
  return `${providerId}:${state}`;
}

function pruneOAuthCallbackCaptures(now = Date.now()) {
  for (const [key, capture] of oauthCallbackCaptures.entries()) {
    if (now - capture.receivedAt > OAUTH_CALLBACK_CAPTURE_TTL_MS) {
      oauthCallbackCaptures.delete(key);
    }
  }

  while (oauthCallbackCaptures.size > MAX_OAUTH_CALLBACK_CAPTURES) {
    const oldestKey = oauthCallbackCaptures.keys().next().value;
    if (typeof oldestKey !== "string") break;
    oauthCallbackCaptures.delete(oldestKey);
  }
}

function isSafeOAuthState(state) {
  return typeof state === "string"
    && state.length > 0
    && state.length <= 512
    && /^[A-Za-z0-9._~-]+$/.test(state);
}

function htmlEscape(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function sendJson(res, statusCode, payload) {
  res.statusCode = statusCode;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(payload));
}

function storeOAuthCallbackCapture(config, callbackUrl) {
  const code = callbackUrl.searchParams.get("code");
  const state = callbackUrl.searchParams.get("state");

  if (!code || !state || !isSafeOAuthState(state)) {
    return null;
  }

  pruneOAuthCallbackCaptures();

  const capture = {
    providerId: config.providerId,
    code,
    state,
    url: callbackUrl.toString(),
    receivedAt: Date.now(),
  };

  oauthCallbackCaptures.set(oauthCallbackKey(capture.providerId, state), capture);
  return capture;
}

function buildOAuthCallbackHtml(config, capture) {
  const fallbackUrl = capture?.url || "";
  const title = capture ? `${config.label} login captured` : `${config.label} login callback was incomplete`;
  const lead = capture
    ? "You can return to Pi for Excel. The add-in should continue automatically in a moment."
    : "Pi for Excel could not find an authorization code in this callback. Return to Pi for Excel and try logging in again.";
  const closeHint = "This browser tab can be closed.";
  const fallbackBlock = capture
    ? `<details><summary>If Pi for Excel did not continue automatically</summary><p>Copy this callback URL and paste it into Pi for Excel:</p><code>${htmlEscape(fallbackUrl)}</code></details>`
    : "";

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pi for Excel login captured</title>
<style>
  body { margin: 0; min-height: 100vh; display: grid; place-items: center; background: #f8faf8; color: #17211b; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
  main { width: min(560px, calc(100vw - 32px)); padding: 28px; border: 1px solid #dfe7df; border-radius: 18px; background: white; box-shadow: 0 18px 55px rgba(15, 23, 18, 0.12); }
  h1 { margin: 0 0 10px; font-size: 22px; }
  p { margin: 0 0 14px; line-height: 1.5; color: #425148; }
  details { margin-top: 18px; }
  summary { cursor: pointer; color: #1d6f42; font-weight: 600; }
  code { display: block; margin-top: 10px; padding: 10px; overflow-wrap: anywhere; border-radius: 10px; background: #f3f6f3; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 12px; }
</style>
</head>
<body>
<main>
  <h1>${htmlEscape(title)}</h1>
  <p>${htmlEscape(lead)}</p>
  <p>${htmlEscape(closeHint)}</p>
  ${fallbackBlock}
</main>
<script>
  try { window.history.replaceState(null, "", "${config.path}/complete"); } catch (_) {}
</script>
</body>
</html>`;
}

function handleOAuthProviderCallbackRequest(configs, port, req, res) {
  const remote = req.socket?.remoteAddress;
  if (!isAllowedClientAddress(remote, [])) {
    res.statusCode = 403;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("forbidden");
    console.warn(`[oauth-callback] blocked non-loopback client: ${remote || "unknown"}`);
    return;
  }

  const rawUrl = req.url || "/";
  let callbackUrl;
  try {
    callbackUrl = new URL(rawUrl, `http://${OAUTH_CALLBACK_HOST}:${port}`);
  } catch {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Invalid callback URL");
    return;
  }

  const config = configs.find((candidate) => candidate.path === callbackUrl.pathname);
  if (!config) {
    res.statusCode = 404;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Not found");
    return;
  }

  const capture = storeOAuthCallbackCapture(config, callbackUrl);
  if (!capture) {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.end(buildOAuthCallbackHtml(config, null));
    return;
  }

  res.statusCode = 200;
  res.setHeader("Content-Type", "text/html; charset=utf-8");
  res.end(buildOAuthCallbackHtml(config, capture));
}

function isOAuthCallbackApiPath(pathname) {
  return pathname.startsWith("/oauth/callback/");
}

function handleOAuthCallbackApiRequest(rawUrl, res) {
  let requestUrl;
  try {
    requestUrl = new URL(rawUrl, "http://proxy.local");
  } catch {
    sendJson(res, 400, { status: "error", error: "invalid_request_url" });
    return;
  }

  let providerId;
  try {
    providerId = decodeURIComponent(requestUrl.pathname.slice("/oauth/callback/".length));
  } catch {
    sendJson(res, 400, { status: "error", error: "invalid_provider" });
    return;
  }

  if (!findOAuthCallbackProvider(providerId)) {
    sendJson(res, 404, { status: "error", error: "unsupported_provider" });
    return;
  }

  const state = requestUrl.searchParams.get("state") || "";
  if (!isSafeOAuthState(state)) {
    sendJson(res, 400, { status: "error", error: "invalid_state" });
    return;
  }

  pruneOAuthCallbackCaptures();

  const capture = oauthCallbackCaptures.get(oauthCallbackKey(providerId, state));
  if (!capture) {
    sendJson(res, 200, { status: "pending" });
    return;
  }

  sendJson(res, 200, {
    status: "ready",
    providerId: capture.providerId,
    code: capture.code,
    state: capture.state,
    url: capture.url,
    receivedAt: capture.receivedAt,
  });
}

function extractProxyTransport(rawUrl) {
  try {
    return new URL(rawUrl, "http://proxy.local").searchParams.get("pi_transport");
  } catch {
    return null;
  }
}

function extractTargetUrl(rawUrl) {
  // rawUrl looks like: /?url=https%3A%2F%2Fapi.example.com/path
  // NOTE: some callers append path segments after the encoded baseUrl,
  // so we decode everything after `url=` rather than using URLSearchParams.
  const idx = rawUrl.indexOf("url=");
  if (idx === -1) return null;
  const encoded = rawUrl.slice(idx + 4);
  const normalized = encoded.replace(/\+/g, "%20");
  try {
    return decodeURIComponent(normalized);
  } catch {
    return null;
  }
}

function buildOutboundHeaders(inHeaders) {
  const out = new Headers();
  for (const [key, value] of Object.entries(inHeaders)) {
    if (!value) continue;
    const lower = key.toLowerCase();

    if (lower === "host") continue;
    if (lower === "content-length") continue;
    if (lower === "accept-encoding") continue;
    if (lower === "user-agent") continue;
    if (lower === "accept-language") continue;

    // Strip browser-only / CORS-triggering headers (mimic server requests)
    if (lower === "origin") continue;
    if (lower === "referer") continue;
    if (lower.startsWith("sec-fetch-")) continue;
    if (lower.startsWith("sec-ch-")) continue;

    // Anthropic uses this header to explicitly enable direct browser access.
    // When proxying we want the upstream to behave like a server-to-server call.
    if (lower === "anthropic-dangerous-direct-browser-access") continue;

    // Never forward cookies through a generic proxy
    if (lower === "cookie") continue;

    if (HOP_BY_HOP_HEADERS.has(lower)) continue;

    if (Array.isArray(value)) {
      for (const v of value) out.append(key, v);
    } else {
      out.set(key, value);
    }
  }
  return out;
}

const handler = async (req, res) => {
  const remote = req.socket?.remoteAddress;
  if (!isAllowedClientAddress(remote, allowedClientCidrs)) {
    res.statusCode = 403;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("forbidden");
    console.warn(`[proxy] blocked disallowed client: ${remote || "unknown"}`);
    return;
  }

  // Health endpoint for load balancers / monitoring. Sits after the client
  // address check but before origin enforcement (health checks often send no
  // Origin header). Browser-based status probes still need CORS headers when
  // they come from an allowed add-in origin. Never proxies anything.
  if ((req.url || "").split("?")[0] === "/healthz") {
    setCorsHeaders(req, res);
    if (req.method === "OPTIONS") {
      res.statusCode = 204;
      res.end();
      return;
    }
    res.statusCode = 200;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.setHeader("X-Pi-For-Excel-Proxy", "1");
    res.setHeader(CODEX_WEBSOCKET_BRIDGE_HEADER, "1");
    res.end("ok");
    return;
  }

  const origin = req.headers.origin;
  if (!isAllowedOrigin(origin)) {
    res.statusCode = 403;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("forbidden");
    console.warn(`[proxy] blocked request from disallowed origin: ${origin || "(none)"}`);
    return;
  }

  setCorsHeaders(req, res);

  if (req.method === "OPTIONS") {
    res.statusCode = 204;
    res.end();
    return;
  }

  const rawUrl = req.url || "/";
  const requestPathname = rawUrl.split("?")[0] || "/";
  if (isOAuthCallbackApiPath(requestPathname)) {
    handleOAuthCallbackApiRequest(rawUrl, res);
    return;
  }

  const target = extractTargetUrl(rawUrl);
  if (!target) {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Missing or invalid ?url=<target-url> query parameter");
    return;
  }

  let targetUrl;
  try {
    targetUrl = new URL(target);
  } catch {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Invalid target URL");
    return;
  }

  if (targetUrl.protocol !== "http:" && targetUrl.protocol !== "https:") {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Only http(s) target URLs are supported");
    return;
  }

  const targetHost = normalizeHost(targetUrl.hostname);
  const safeTarget = `${targetUrl.origin}${targetUrl.pathname}`;

  let resolvedIps = [];
  if (!isIpLiteral(targetHost)) {
    try {
      const records = await dnsLookup(targetHost, { all: true, verbatim: true });
      resolvedIps = records.map((r) => r.address);
    } catch (err) {
      const errorText = err instanceof Error ? err.message : String(err);
      if (strictTargetResolution) {
        rejectWithReason(res, "blocked_target_resolution_failed");
        console.warn(`[proxy] blocked target (blocked_target_resolution_failed): ${safeTarget} (${errorText})`);
        return;
      }
      console.warn(`[proxy] DNS lookup failed for ${targetHost}: ${errorText} (continuing)`);
    }
  }

  const bypassHostAllowlistForGitHubEnterprise =
    !allowAllTargetHosts
    && configuredAllowedTargetHosts.size === 0
    && shouldBypassHostAllowlistForGitHubEnterprise(targetUrl);

  const effectiveAllowedTargetHosts = bypassHostAllowlistForGitHubEnterprise
    ? EMPTY_ALLOWED_TARGET_HOSTS
    : allowedTargetHosts;

  const targetPolicy = evaluateTargetHostPolicy({
    hostname: targetHost,
    resolvedIps,
    allowLoopbackTargets,
    allowPrivateTargets,
    allowedHosts: effectiveAllowedTargetHosts,
    // SECURITY: when the operator explicitly configured ALLOWED_TARGET_HOSTS,
    // loopback/private override flags must not bypass it (central proxies).
    requireAllowlistForOverriddenTargets: configuredAllowedTargetHosts.size > 0,
  });

  if (!targetPolicy.allowed) {
    const reason = targetPolicy.reason || "forbidden";
    rejectWithReason(res, reason);
    console.warn(`[proxy] blocked target (${reason}): ${safeTarget}`);
    return;
  }

  if (bypassHostAllowlistForGitHubEnterprise) {
    console.log(`[proxy] allowing GitHub enterprise endpoint outside default host allowlist: ${safeTarget}`);
  }

  const requestedTransport = extractProxyTransport(rawUrl);
  if (requestedTransport && requestedTransport !== CODEX_WEBSOCKET_BRIDGE_TRANSPORT) {
    res.statusCode = 400;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("Unsupported pi_transport value");
    return;
  }

  if (requestedTransport === CODEX_WEBSOCKET_BRIDGE_TRANSPORT) {
    if (!isCodexWebSocketBridgeTarget(targetUrl)) {
      res.statusCode = 400;
      res.setHeader("Content-Type", "text/plain; charset=utf-8");
      res.end("Codex WebSocket bridge target must be https://chatgpt.com/backend-api/codex/responses");
      return;
    }

    const startedAt = Date.now();
    const headers = buildOutboundHeaders(req.headers);
    await bridgeCodexWebSocketToSse({ req, res, targetUrl, outboundHeaders: headers });
    console.log(`[proxy] ${req.method || "GET"} ${safeTarget} via Codex WebSocket bridge (${Date.now() - startedAt}ms)`);
    return;
  }

  try {
    const startedAt = Date.now();
    const headers = buildOutboundHeaders(req.headers);

    const hasBody = req.method && !["GET", "HEAD"].includes(req.method);
    const body = hasBody ? Readable.toWeb(req) : undefined;

    const upstream = await fetch(targetUrl.toString(), {
      method: req.method,
      headers,
      body,
      // Required when using a stream body in Node fetch
      ...(body ? { duplex: "half" } : {}),
      redirect: "manual",
    });

    // Log without query string to avoid leaking tokens
    console.log(`[proxy] ${req.method || "GET"} ${safeTarget} -> ${upstream.status} (${Date.now() - startedAt}ms)`);

    res.statusCode = upstream.status;

    // Copy response headers
    upstream.headers.forEach((value, key) => {
      const lower = key.toLowerCase();
      if (lower === "set-cookie") return;
      if (HOP_BY_HOP_HEADERS.has(lower)) return;
      // Node fetch transparently decompresses responses but keeps the original
      // Content-Encoding header (e.g. "gzip"). Forwarding that header would
      // make the browser try to decompress *again* and fail while reading.
      if (lower === "content-encoding") return;

      // Content-Length can be wrong after decompression; let Node set it.
      if (lower === "content-length") return;

      // Keep our CORS headers (set by setCorsHeaders). Upstream values could
      // clobber them and break the integration — e.g. llama.cpp returns an
      // empty Access-Control-Allow-Origin because we don't forward the Origin.
      if (lower.startsWith("access-control-")) return;
      if (lower === "vary") return;

      res.setHeader(key, value);
    });

    if (!upstream.body) {
      res.end();
      return;
    }

    const nodeStream = Readable.fromWeb(upstream.body);
    nodeStream.on("error", () => {
      try {
        res.end();
      } catch {
        // ignore
      }
    });
    nodeStream.pipe(res);
  } catch (err) {
    console.warn(`[proxy] ${req.method || "GET"} ${targetUrl.origin}${targetUrl.pathname} -> ERROR (${err instanceof Error ? err.message : String(err)})`);
    res.statusCode = 502;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end(`Proxy error: ${err instanceof Error ? err.message : String(err)}`);
  }
};

function createProxyServer() {
  if (!useHttps) {
    return http.createServer(handler);
  }

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error("[pi-for-excel] HTTPS requested but TLS key/cert not found:");
    console.error(`  key:  ${keyPath}${fs.existsSync(keyPath) ? "" : "  (missing)"}`);
    console.error(`  cert: ${certPath}${fs.existsSync(certPath) ? "" : "  (missing)"}`);
    console.error("For local dev, generate key.pem/cert.pem with mkcert (see README). Example: mkcert localhost");
    console.error("For central deployments, set TLS_KEY_PATH/TLS_CERT_PATH (see docs/central-proxy.md).");
    process.exit(1);
  }

  return https.createServer(
    {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath),
    },
    handler,
  );
}

function logStartup(listeningPort, listeningHosts) {
  const scheme = useHttps ? "https" : "http";
  const formattedHost = HOST.includes(":") && !HOST.startsWith("[") ? `[${HOST}]` : HOST;
  const proxyUrl = `${scheme}://${formattedHost}:${listeningPort}`;
  console.log(`[pi-for-excel] CORS proxy listening on ${proxyUrl}`);
  if (listeningHosts.length > 1) {
    console.log(`[pi-for-excel] Listening on loopback addresses: ${listeningHosts.join(", ")}`);
  }
  console.log(`[pi-for-excel] Format: ${proxyUrl}/?url=<target-url>`);
  if (listeningPort !== DEFAULT_PORT) {
    console.log(`[pi-for-excel] Update Pi for Excel /settings → Proxy URL to ${proxyUrl}`);
  }
  console.log(`[pi-for-excel] Allowed origins: ${Array.from(allowedOrigins).join(", ")}`);

  if (allowedClientCidrs.length > 0) {
    console.log(`[pi-for-excel] WARNING: accepting non-loopback clients from: ${allowedClientCidrs.map((c) => c.entry).join(", ")} (ALLOWED_CLIENT_CIDRS)`);
    console.log("[pi-for-excel] Ensure network-level controls also restrict who can reach this proxy.");
  } else {
    console.log("[pi-for-excel] Client policy: loopback only");
  }

  if (allowAllTargetHosts) {
    console.log("[pi-for-excel] WARNING: target host allowlisting disabled (ALLOW_ALL_TARGET_HOSTS=1)");
  } else {
    const source = configuredAllowedTargetHosts.size > 0 ? "ALLOWED_TARGET_HOSTS" : "default";
    console.log(`[pi-for-excel] Allowed target hosts (${source}): ${Array.from(allowedTargetHosts).join(", ")}`);

    if (configuredAllowedTargetHosts.size === 0) {
      console.log("[pi-for-excel] GitHub enterprise OAuth/Copilot endpoints on custom domains are allowed by path.");
    }
  }

  if (allowLoopbackTargets) {
    console.log("[pi-for-excel] WARNING: loopback target blocking disabled (ALLOW_LOOPBACK_TARGETS=1)");
  }

  if (allowPrivateTargets) {
    console.log("[pi-for-excel] WARNING: private/local target blocking disabled (ALLOW_PRIVATE_TARGETS=1)");
  }

  if (strictTargetResolution) {
    console.log("[pi-for-excel] Strict DNS resolution enabled (STRICT_TARGET_RESOLUTION=1)");
  }
}

function groupOAuthCallbackProvidersByPort() {
  const byPort = new Map();
  for (const config of OAUTH_CALLBACK_PROVIDER_CONFIGS) {
    const existing = byPort.get(config.port);
    if (existing) {
      existing.push(config);
    } else {
      byPort.set(config.port, [config]);
    }
  }
  return byPort;
}

const oauthCallbackServers = OAUTH_CALLBACK_SERVER_ENABLED
  ? Array.from(groupOAuthCallbackProvidersByPort(), ([port, configs]) => ({
      port,
      configs,
      server: http.createServer((req, res) => handleOAuthProviderCallbackRequest(configs, port, req, res)),
    }))
  : [];

function listenOAuthCallbackServer(entry) {
  let cleanedUp = false;
  const cleanup = () => {
    if (cleanedUp) return;
    cleanedUp = true;
    entry.server.off("error", onError);
    entry.server.off("listening", onListening);
  };

  const onError = (err) => {
    cleanup();
    const code = typeof err?.code === "string" ? err.code : "listen_error";
    console.warn(`[pi-for-excel] OAuth callback capture listener unavailable (${code}).`);
    console.warn("[pi-for-excel] Affected OAuth logins will still work with manual callback URL paste.");
  };

  const onListening = () => {
    cleanup();
    console.log("[pi-for-excel] OAuth callback capture listener started.");
  };

  entry.server.once("error", onError);
  entry.server.once("listening", onListening);
  entry.server.listen(entry.port, OAUTH_CALLBACK_HOST);
}

function listenOAuthCallbackServers() {
  for (const entry of oauthCallbackServers) {
    listenOAuthCallbackServer(entry);
  }
}

function closeServer(server) {
  return new Promise((resolve) => {
    try {
      server.close(() => resolve());
    } catch {
      resolve();
    }
  });
}

function listenServer(server, port, host) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const cleanup = () => {
      server.off("error", onError);
      server.off("listening", onListening);
    };

    const onError = (error) => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(error);
    };

    const onListening = () => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve();
    };

    server.once("error", onError);
    server.once("listening", onListening);
    server.listen(port, host);
  });
}

async function listen(port) {
  const entries = [];
  const skippedHosts = [];
  let selectedPort = port;

  try {
    for (const host of LISTEN_HOSTS) {
      const server = createProxyServer();
      try {
        await listenServer(server, selectedPort, host);
      } catch (error) {
        await closeServer(server);
        const code = typeof error?.code === "string" ? error.code : "";
        if (!hasExplicitHost && (code === "EAFNOSUPPORT" || code === "EADDRNOTAVAIL")) {
          skippedHosts.push(host);
          continue;
        }
        await Promise.all(entries.map((entry) => closeServer(entry.server)));
        throw error;
      }

      const address = server.address();
      const actualPort = address && typeof address !== "string" ? address.port : selectedPort;
      if (selectedPort === 0) {
        selectedPort = actualPort;
      }
      entries.push({ server, host, port: actualPort });
    }

    if (entries.length === 0) {
      throw new Error(
        skippedHosts.length > 0
          ? `No loopback listen addresses were available (${skippedHosts.join(", ")})`
          : "No listen addresses were configured",
      );
    }

    if (skippedHosts.length > 0) {
      console.warn(`[pi-for-excel] Skipped unavailable loopback addresses: ${skippedHosts.join(", ")}`);
    }
    logStartup(selectedPort, entries.map((entry) => entry.host));
  } catch (error) {
    const code = typeof error?.code === "string" ? error.code : "";
    if (code === "EADDRINUSE" && !hasExplicitPort && port === DEFAULT_PORT) {
      console.warn(`[pi-for-excel] Port ${DEFAULT_PORT} is already in use; choosing a random available port instead.`);
      await listen(0);
      return;
    }

    const message = error instanceof Error ? error.message : String(error);
    console.error(`[pi-for-excel] Failed to listen on ${HOST}:${port}: ${message}`);
    if (hasExplicitPort) {
      console.error("[pi-for-excel] Choose a different port with PORT=0 (random) or PORT=<port>.");
    }
    process.exit(1);
  }
}

await listen(PORT);
listenOAuthCallbackServers();
