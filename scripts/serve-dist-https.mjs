#!/usr/bin/env node

/**
 * Serve the Vite production build (dist/) over HTTPS.
 *
 * Useful for Office add-in smoke tests where the manifest must load from HTTPS
 * but we want to exercise the *production* code path (import.meta.env.DEV === false).
 *
 * Requires mkcert-generated certs in the repo root:
 *   - cert.pem
 *   - key.pem
 */

import https from "node:https";
import fs from "node:fs";
import path from "node:path";

const HOST = process.env.HOST || "localhost";
const PORT = Number.parseInt(process.env.PORT || "3000", 10);

const rootDir = path.resolve(process.cwd());
const distDir = path.join(rootDir, "dist");

const keyPath = path.join(rootDir, "key.pem");
const certPath = path.join(rootDir, "cert.pem");

if (!fs.existsSync(distDir)) {
  console.error(`[serve-dist-https] dist/ not found at ${distDir}. Run: npm run build`);
  process.exit(1);
}

if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
  console.error("[serve-dist-https] Missing HTTPS certs (key.pem/cert.pem) in repo root.");
  console.error("Generate them with mkcert (see README Quick Start). ");
  process.exit(1);
}

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
  ".woff": "font/woff",
  ".ttf": "font/ttf",
  ".txt": "text/plain; charset=utf-8",
};

function safeJoin(base, reqPath) {
  const decoded = decodeURIComponent(reqPath);
  const cleaned = decoded.replace(/^\/+/, "");
  const full = path.resolve(base, cleaned);
  if (!full.startsWith(base)) {
    throw new Error("Path traversal");
  }
  return full;
}

const server = https.createServer(
  {
    key: fs.readFileSync(keyPath),
    cert: fs.readFileSync(certPath),
  },
  (req, res) => {
    try {
      const url = new URL(req.url || "/", `https://${HOST}:${PORT}`);
      let reqPath = url.pathname;
      if (reqPath === "/") reqPath = "/src/taskpane.html";

      const filePath = safeJoin(distDir, reqPath);
      if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
        res.statusCode = 404;
        res.setHeader("content-type", "text/plain; charset=utf-8");
        console.log(`[serve-dist-https] ${req.method || "GET"} ${reqPath} -> 404`);
        res.end("not found");
        return;
      }

      const ext = path.extname(filePath).toLowerCase();
      res.statusCode = 200;
      res.setHeader("content-type", MIME[ext] || "application/octet-stream");
      console.log(`[serve-dist-https] ${req.method || "GET"} ${reqPath} -> 200`);

      // Cache hashed assets aggressively, keep HTML uncached
      if (reqPath.startsWith("/assets/") && /-[A-Za-z0-9]{8,}\./.test(reqPath)) {
        res.setHeader("cache-control", "public, max-age=31536000, immutable");
      } else {
        res.setHeader("cache-control", "no-cache");
      }

      fs.createReadStream(filePath).pipe(res);
    } catch (e) {
      res.statusCode = 500;
      res.setHeader("content-type", "text/plain; charset=utf-8");
      console.log(`[serve-dist-https] ${req.method || "GET"} ${req.url || "/"} -> 500`);
      res.end(`server error: ${e instanceof Error ? e.message : String(e)}`);
    }
  },
);

// Bind to all interfaces so both IPv4 (127.0.0.1) and IPv6 (::1) localhost
// resolutions work in different Office webviews.
server.listen(PORT, () => {
  console.log(`[serve-dist-https] Serving ${distDir}`);
  console.log(`[serve-dist-https] https://${HOST}:${PORT}/src/taskpane.html`);
});
