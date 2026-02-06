#!/usr/bin/env node

/**
 * Minimal CORS proxy for Pi for Excel.
 *
 * Why this exists:
 * - Some provider OAuth/token endpoints (and some LLM APIs) block browser requests via CORS.
 * - In dev we rely on Vite's proxy. In production, you can run this locally and point
 *   Pi for Excel's proxy setting at it (default: http://localhost:3001).
 *
 * Proxy format:
 *   http://localhost:3001/?url=<target-url>
 *
 * Example:
 *   curl 'http://localhost:3001/?url=https%3A%2F%2Fexample.com' 
 */

import http from "node:http";
import { Readable } from "node:stream";

const HOST = process.env.HOST || "127.0.0.1";
const PORT = Number.parseInt(process.env.PORT || "3001", 10);

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

function setCorsHeaders(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,PATCH,DELETE,OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    req.headers["access-control-request-headers"] || "*",
  );
  res.setHeader("Access-Control-Expose-Headers", "*");
  res.setHeader("Access-Control-Max-Age", "86400");
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

const server = http.createServer(async (req, res) => {
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

  try {
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
    res.statusCode = 502;
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.end(`Proxy error: ${err instanceof Error ? err.message : String(err)}`);
  }
});

server.listen(PORT, HOST, () => {
  console.log(`[pi-for-excel] CORS proxy listening on http://${HOST}:${PORT}`);
  console.log(`[pi-for-excel] Format: http://${HOST}:${PORT}/?url=<target-url>`);
});
