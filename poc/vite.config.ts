import { defineConfig } from "vite";
import fs from "fs";

export default defineConfig({
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
