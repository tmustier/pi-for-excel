import { defineConfig, type Plugin } from "vite";
import fs from "fs";
import path from "path";
import os from "os";

// ============================================================================
// Plugins
// ============================================================================

/**
 * Serves pi's ~/.pi/agent/auth.json so the browser can reuse
 * existing OAuth/API key credentials without re-logging in.
 * Dev-only convenience — production uses its own auth flow.
 */
function piAuthPlugin(): Plugin {
  const authPath = path.join(os.homedir(), ".pi", "agent", "auth.json");
  return {
    name: "pi-auth",
    configureServer(server) {
      server.middlewares.use("/__pi-auth", (_req, res) => {
        try {
          const data = fs.readFileSync(authPath, "utf-8");
          res.setHeader("Content-Type", "application/json");
          res.end(data);
        } catch {
          res.statusCode = 404;
          res.end(JSON.stringify({ error: "auth.json not found" }));
        }
      });
    },
  };
}

// ============================================================================
// Proxy helper — strips browser headers so APIs don't treat requests as CORS
// ============================================================================

/** Common proxy config: strip Origin/Referer so the target sees a server request */
type ProxyReqLike = { removeHeader(name: string): void };
type ProxyServerLike = { on(event: "proxyReq", handler: (proxyReq: ProxyReqLike) => void): void };

function stripBrowserHeaders(proxy: ProxyServerLike) {
  proxy.on("proxyReq", (proxyReq) => {
    proxyReq.removeHeader("origin");
    proxyReq.removeHeader("referer");
    proxyReq.removeHeader("sec-fetch-mode");
    proxyReq.removeHeader("sec-fetch-site");
    proxyReq.removeHeader("sec-fetch-dest");
    proxyReq.removeHeader("anthropic-dangerous-direct-browser-access");
  });
}

function proxyEntry(target: string, proxyPath: string) {
  return {
    target,
    changeOrigin: true,
    rewrite: (p: string) => p.replace(new RegExp(`^${proxyPath.replace(/\//g, "\\/")}`), ""),
    secure: true,
    configure: stripBrowserHeaders,
  };
}

// ============================================================================
// Vite config
// ============================================================================

// HTTPS certs — generate with: mkcert localhost
const keyPath = path.resolve(__dirname, "key.pem");
const certPath = path.resolve(__dirname, "cert.pem");

const hasHttpsCerts = fs.existsSync(keyPath) && fs.existsSync(certPath);

export default defineConfig({
  plugins: [piAuthPlugin()],

  server: {
    port: 3000,
    https: hasHttpsCerts
      ? { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) }
      : undefined,

    proxy: {
      // OAuth token endpoints
      "/oauth-proxy/anthropic": proxyEntry("https://console.anthropic.com", "/oauth-proxy/anthropic"),
      "/oauth-proxy/github": proxyEntry("https://github.com", "/oauth-proxy/github"),

      // API proxies (providers that block browser CORS)
      "/api-proxy/anthropic": proxyEntry("https://api.anthropic.com", "/api-proxy/anthropic"),
      "/api-proxy/openai-auth": proxyEntry("https://auth.openai.com", "/api-proxy/openai-auth"),
      "/api-proxy/openai": proxyEntry("https://api.openai.com", "/api-proxy/openai"),
      "/api-proxy/chatgpt": proxyEntry("https://chatgpt.com", "/api-proxy/chatgpt"),
      "/api-proxy/google-oauth": proxyEntry("https://oauth2.googleapis.com", "/api-proxy/google-oauth"),
      "/api-proxy/google": proxyEntry("https://generativelanguage.googleapis.com", "/api-proxy/google"),
    },
  },

  // Prevent esbuild from downleveling class fields (breaks Lit's @state/@property)
  optimizeDeps: {
    esbuildOptions: { target: "esnext" },
  },
  esbuild: { target: "esnext" },

  // Stub Node.js built-ins imported by Anthropic SDK's transitive deps (undici, @smithy).
  // These code paths are never executed in the browser — all API calls use fetch().
  resolve: {
    alias: {
      stream: path.resolve(__dirname, "src/stubs/stream.ts"),
    },
  },

  build: {
    target: "esnext",
    commonjsOptions: {
      // Ignore node built-in imports that can't be resolved
      ignoreDynamicRequires: true,
    },
    rollupOptions: {
      input: {
        taskpane: "src/taskpane.html",
      },
      // Externalize Node.js-only packages that are never used in browser
      external: [
        /^node:/,
        /^@smithy\//,
      ],
      output: {
        // Map externals to empty modules at runtime
        globals: {
          http: "{}",
          https: "{}",
          net: "{}",
          tls: "{}",
        },
      },
    },
  },
});
