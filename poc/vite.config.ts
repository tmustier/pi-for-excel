import { defineConfig, type Plugin } from "vite";
import fs from "fs";
import path from "path";
import os from "os";

/**
 * Vite plugin that serves pi's auth.json so the browser can reuse
 * existing OAuth/API key credentials without re-logging in.
 * Dev-only — in production the add-in uses its own auth flow.
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

export default defineConfig({
  plugins: [piAuthPlugin()],
  server: {
    port: 3000,
    https: {
      key: fs.readFileSync("key.pem"),
      cert: fs.readFileSync("cert.pem"),
    },
    // Proxy endpoints that don't support browser CORS
    proxy: {
      "/oauth-proxy/anthropic": {
        target: "https://console.anthropic.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/oauth-proxy\/anthropic/, ""),
        secure: true,
        configure: (proxy) => {
          proxy.on("proxyReq", (proxyReq) => {
            proxyReq.removeHeader("origin");
            proxyReq.removeHeader("referer");
          });
        },
      },
      "/oauth-proxy/github": {
        target: "https://github.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/oauth-proxy\/github/, ""),
        secure: true,
      },
      // OpenAI auth (token refresh)
      "/api-proxy/openai-auth": {
        target: "https://auth.openai.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api-proxy\/openai-auth/, ""),
        secure: true,
        configure: (proxy) => {
          proxy.on("proxyReq", (proxyReq) => {
            proxyReq.removeHeader("origin");
            proxyReq.removeHeader("referer");
          });
        },
      },
      // OpenAI API
      "/api-proxy/openai": {
        target: "https://api.openai.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api-proxy\/openai/, ""),
        secure: true,
        configure: (proxy) => {
          proxy.on("proxyReq", (proxyReq) => {
            proxyReq.removeHeader("origin");
            proxyReq.removeHeader("referer");
          });
        },
      },
      // Google OAuth (token refresh)
      "/api-proxy/google-oauth": {
        target: "https://oauth2.googleapis.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api-proxy\/google-oauth/, ""),
        secure: true,
        configure: (proxy) => {
          proxy.on("proxyReq", (proxyReq) => {
            proxyReq.removeHeader("origin");
            proxyReq.removeHeader("referer");
          });
        },
      },
      // Google Generative AI API
      "/api-proxy/google": {
        target: "https://generativelanguage.googleapis.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api-proxy\/google/, ""),
        secure: true,
        configure: (proxy) => {
          proxy.on("proxyReq", (proxyReq) => {
            proxyReq.removeHeader("origin");
            proxyReq.removeHeader("referer");
          });
        },
      },
      // OpenAI Codex API (chatgpt.com/backend-api — NOT api.openai.com)
      "/api-proxy/chatgpt": {
        target: "https://chatgpt.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api-proxy\/chatgpt/, ""),
        secure: true,
        configure: (proxy) => {
          proxy.on("proxyReq", (proxyReq) => {
            proxyReq.removeHeader("origin");
            proxyReq.removeHeader("referer");
            proxyReq.removeHeader("sec-fetch-mode");
            proxyReq.removeHeader("sec-fetch-site");
            proxyReq.removeHeader("sec-fetch-dest");
          });
        },
      },
      // Anthropic API - CORS blocked for OAuth org tokens
      "/api-proxy/anthropic": {
        target: "https://api.anthropic.com",
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api-proxy\/anthropic/, ""),
        secure: true,
        configure: (proxy) => {
          // Strip browser-identifying headers so Anthropic doesn't treat this as CORS
          proxy.on("proxyReq", (proxyReq) => {
            proxyReq.removeHeader("origin");
            proxyReq.removeHeader("referer");
            proxyReq.removeHeader("sec-fetch-mode");
            proxyReq.removeHeader("sec-fetch-site");
            proxyReq.removeHeader("sec-fetch-dest");
            // Also remove the dangerous header if still present
            proxyReq.removeHeader("anthropic-dangerous-direct-browser-access");
          });
        },
      },
    },
  },
  // Fix Lit class field shadowing: esbuild downlevels native class fields to
  // __publicField() calls that create own properties, shadowing Lit's @state()
  // prototype accessors. Setting target to esnext prevents this transformation.
  // See: https://lit.dev/msg/class-field-shadowing
  optimizeDeps: {
    esbuildOptions: {
      target: "esnext",
    },
  },
  esbuild: {
    target: "esnext",
  },
  build: {
    target: "esnext",
    rollupOptions: {
      input: {
        taskpane: "src/taskpane.html",
      },
    },
  },
});
