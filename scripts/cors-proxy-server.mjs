#!/usr/bin/env node

/**
 * Minimal CORS proxy for Pi for Excel.
 *
 * Why this exists:
 * - Some provider OAuth/token endpoints (and some LLM APIs) block browser requests via CORS.
 * - In dev we rely on Vite's proxy. In production, you can run this locally and point
 *   Pi for Excel's proxy setting at it (default: https://localhost:3003).
 *
 * Usage:
 *   npm run proxy:https   # HTTPS (recommended for Office webviews)
 *   npm run proxy         # HTTP  (may be blocked as mixed content)
 *
 * Proxy format:
 *   https://localhost:3003/?url=<target-url>
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
  evaluateTargetHostPolicy,
  isIpLiteral,
  normalizeHost,
  parseAllowedTargetHosts,
} from "./proxy-target-policy.mjs";

const args = new Set(process.argv.slice(2));
const useHttps = args.has("--https") || process.env.HTTPS === "1" || process.env.HTTPS === "true";
const useHttp = args.has("--http");

if (useHttps && useHttp) {
  console.error("[pi-for-excel] Invalid args: can't use both --https and --http");
  process.exit(1);
}

const HOST = process.env.HOST || (useHttps ? "localhost" : "127.0.0.1");
const PORT = Number.parseInt(process.env.PORT || "3003", 10);

const rootDir = path.resolve(process.cwd());
const keyPath = path.join(rootDir, "key.pem");
const certPath = path.join(rootDir, "cert.pem");

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
  "https://localhost:3000",
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

function isLoopbackAddress(addr) {
  if (!addr) return false;
  if (addr === "::1" || addr === "0:0:0:0:0:0:0:1") return true;
  if (addr.startsWith("127.")) return true;
  if (addr.startsWith("::ffff:127.")) return true;
  return false;
}

function envFlag(name) {
  const raw = process.env[name];
  return raw === "1" || raw === "true";
}

const DEFAULT_ALLOWED_TARGET_HOSTS = new Set([
  "api.anthropic.com",
  "console.anthropic.com",
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
  res.setHeader("Access-Control-Expose-Headers", "*");
  res.setHeader("Access-Control-Max-Age", "86400");
}

function rejectWithReason(res, reason) {
  const msg = TARGET_POLICY_MESSAGES[reason] || "forbidden";
  res.statusCode = 403;
  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.end(`${reason}: ${msg}`);
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

    // Strip browser-only / CORS-triggering headers (mimic server requests)
    if (lower === "origin") continue;
    if (lower === "referer") continue;
    if (lower.startsWith("sec-fetch-")) continue;

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
  if (!isLoopbackAddress(remote)) {
    res.statusCode = 403;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end("forbidden");
    console.warn(`[proxy] blocked non-loopback client: ${remote || "unknown"}`);
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

    // Copy response headers (but keep our CORS headers)
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

const server = (() => {
  if (!useHttps) {
    return http.createServer(handler);
  }

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error("[pi-for-excel] HTTPS requested but key.pem/cert.pem not found in repo root.");
    console.error("Generate them with mkcert (see README). Example: mkcert localhost");
    process.exit(1);
  }

  return https.createServer(
    {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath),
    },
    handler,
  );
})();

server.listen(PORT, HOST, () => {
  const scheme = useHttps ? "https" : "http";
  console.log(`[pi-for-excel] CORS proxy listening on ${scheme}://${HOST}:${PORT}`);
  console.log(`[pi-for-excel] Format: ${scheme}://${HOST}:${PORT}/?url=<target-url>`);
  console.log(`[pi-for-excel] Allowed origins: ${Array.from(allowedOrigins).join(", ")}`);

  if (allowAllTargetHosts) {
    console.log("[pi-for-excel] WARNING: target host allowlisting disabled (ALLOW_ALL_TARGET_HOSTS=1)");
  } else {
    const source = configuredAllowedTargetHosts.size > 0 ? "ALLOWED_TARGET_HOSTS" : "default";
    console.log(`[pi-for-excel] Allowed target hosts (${source}): ${Array.from(allowedTargetHosts).join(", ")}`);

    if (configuredAllowedTargetHosts.size === 0) {
      console.log("[pi-for-excel] GitHub enterprise OAuth/Copilot endpoints on custom domains are allowed by path.");
    }
  }

  if (hasConfiguredAllowedTargetHosts && configuredAllowedTargetHosts.size === 0) {
    console.warn("[pi-for-excel] WARNING: ALLOWED_TARGET_HOSTS had no valid entries; using default allowlist.");
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
});
