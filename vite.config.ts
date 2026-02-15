import { defineConfig, type Plugin } from "vite";
import type { IncomingMessage, ServerResponse } from "node:http";
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

  const isLoopbackAddress = (addr: string | undefined): boolean => {
    if (!addr) return false;
    if (addr === "::1" || addr === "0:0:0:0:0:0:0:1") return true;
    if (addr.startsWith("127.")) return true;
    if (addr.startsWith("::ffff:127.")) return true;
    return false;
  };

  return {
    name: "pi-auth",
    configureServer(server) {
      server.middlewares.use("/__pi-auth", (req: IncomingMessage, res: ServerResponse) => {
        // SECURITY: auth.json can contain API keys + refresh tokens.
        // Only serve it to loopback clients (Excel webviews, local browser).
        const remote = req.socket?.remoteAddress;
        if (!isLoopbackAddress(remote)) {
          res.statusCode = 403;
          res.setHeader("Content-Type", "application/json; charset=utf-8");
          res.setHeader("Cache-Control", "no-store");
          res.end(JSON.stringify({ error: "forbidden" }));
          return;
        }

        try {
          const data = fs.readFileSync(authPath, "utf-8");
          res.setHeader("Content-Type", "application/json; charset=utf-8");
          res.setHeader("Cache-Control", "no-store");
          res.end(data);
        } catch {
          res.statusCode = 404;
          res.setHeader("Content-Type", "application/json; charset=utf-8");
          res.setHeader("Cache-Control", "no-store");
          res.end(JSON.stringify({ error: "auth.json not found" }));
        }
      });
    },
  };
}

/**
 * Stub out the Amazon Bedrock provider in browser builds.
 *
 * pi-ai registers all built-in providers at import time, including Bedrock.
 * The Bedrock provider pulls in AWS SDK Node transports which break Vite's
 * production bundling for the browser.
 */
function stubBedrockProviderPlugin(): Plugin {
  const stubPath = path.resolve(__dirname, "src/stubs/amazon-bedrock.ts");

  return {
    name: "stub-bedrock-provider",
    enforce: "pre",
    resolveId(id, importer) {
      // Register-builtins imports Bedrock via a relative path.
      if (
        id === "./amazon-bedrock.js" &&
        importer &&
        importer.includes("@mariozechner/pi-ai") &&
        importer.includes("providers/register-builtins")
      ) {
        return stubPath;
      }

      // Safety: also catch resolved imports.
      if (id.includes("@mariozechner/pi-ai") && id.endsWith("/providers/amazon-bedrock.js")) {
        return stubPath;
      }

      return null;
    },
  };
}

/**
 * Stub out pi-ai's OAuth index in browser builds.
 *
 * pi-ai's main entrypoint re-exports the OAuth index, which includes Node-only
 * side effects and CLI-only providers. The Excel add-in uses a small, local
 * OAuth implementation and should not bundle these flows.
 */
function stubPiAiOAuthIndexPlugin(): Plugin {
  const stubPath = path.resolve(__dirname, "src/stubs/pi-ai-oauth.ts");

  return {
    name: "stub-pi-ai-oauth-index",
    enforce: "pre",
    resolveId(id, importer) {
      const cleanId = id.split("?")[0];
      const cleanImporter = importer?.split("?")[0];

      // pi-ai's dist/index.js re-exports the OAuth index via a relative path.
      if (
        cleanId === "./utils/oauth/index.js" &&
        cleanImporter &&
        cleanImporter.includes("/node_modules/@mariozechner/pi-ai/") &&
        cleanImporter.endsWith("/dist/index.js")
      ) {
        return stubPath;
      }

      // Safety: catch resolved ids too.
      if (cleanId.includes("/node_modules/@mariozechner/pi-ai/") && cleanId.endsWith("/dist/utils/oauth/index.js")) {
        return stubPath;
      }

      return null;
    },
  };
}

/**
 * Stub out pi-web-ui tool modules we don't ship in the Excel add-in.
 *
 * pi-web-ui's tools/index.js auto-imports optional tools (document extraction,
 * JavaScript REPL) to register their renderers. Those pull heavy dependencies
 * (pdfjs-dist, docx-preview, xlsx) that bloat the taskpane bundle.
 */
function stubPiWebUiBuiltinToolsPlugin(): Plugin {
  const stubExtractDocumentPath = path.resolve(__dirname, "src/stubs/pi-web-ui-extract-document.ts");
  const stubJavascriptReplPath = path.resolve(__dirname, "src/stubs/pi-web-ui-javascript-repl.ts");
  const stubAttachmentUtilsPath = path.resolve(__dirname, "src/stubs/pi-web-ui-attachment-utils.ts");
  const stubAttachmentOverlayPath = path.resolve(__dirname, "src/stubs/pi-web-ui-attachment-overlay.ts");
  const stubArtifactsPanelPath = path.resolve(__dirname, "src/stubs/pi-web-ui-artifacts-panel.ts");
  const stubArtifactsToolRendererPath = path.resolve(__dirname, "src/stubs/pi-web-ui-artifacts-tool-renderer.ts");

  const norm = (p: string): string => p.split("?")[0].replaceAll("\\", "/");

  return {
    name: "stub-pi-web-ui-builtin-tools",
    enforce: "pre",
    resolveId(id, importer) {
      const cleanId = norm(id);
      const cleanImporter = importer ? norm(importer) : "";

      const importerIsPiWebUi = cleanImporter.includes("@mariozechner/pi-web-ui");
      if (!importerIsPiWebUi) return null;

      // ── Tools (pi-web-ui ships them, but Excel add-in does not) ──
      if (cleanImporter.endsWith("/dist/tools/index.js")) {
        // tools/index.js imports these via relative paths.
        if (cleanId === "./extract-document.js") return stubExtractDocumentPath;
        if (cleanId === "./javascript-repl.js") return stubJavascriptReplPath;
      }

      // index.js re-exports these via ./tools/*
      if (cleanId.endsWith("tools/extract-document.js")) return stubExtractDocumentPath;
      if (cleanId.endsWith("tools/javascript-repl.js")) return stubJavascriptReplPath;

      // ── Attachments (heavy deps: pdfjs-dist, docx-preview, xlsx) ──
      if (cleanId.endsWith("utils/attachment-utils.js")) return stubAttachmentUtilsPath;
      if (cleanId.endsWith("dialogs/AttachmentOverlay.js")) return stubAttachmentOverlayPath;

      // ── Artifacts (pull in PDF/DOCX/XLSX renderers) ──
      if (cleanId.endsWith("tools/artifacts/artifacts.js")) return stubArtifactsPanelPath;
      if (cleanId.endsWith("tools/artifacts/artifacts-tool-renderer.js")) return stubArtifactsToolRendererPath;

      return null;
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

function escapeRegExp(value: string): string {
  return value.replace(/[\\^$.*+?()[\]{}|]/g, "\\$&");
}

function proxyEntry(target: string, proxyPath: string) {
  const escapedProxyPath = escapeRegExp(proxyPath);

  return {
    target,
    changeOrigin: true,
    rewrite: (p: string) => p.replace(new RegExp(`^${escapedProxyPath}`), ""),
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
  plugins: [
    piAuthPlugin(),
    stubBedrockProviderPlugin(),
    stubPiAiOAuthIndexPlugin(),
    stubPiWebUiBuiltinToolsPlugin(),
  ],

  server: {
    // Must stay on :3000 because manifest hardcodes it.
    // Bind IPv6 too: Excel's webview may resolve localhost → ::1 and fail if we only listen on 127.0.0.1.
    host: "::",
    strictPort: true,
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
      "/api-proxy/google-cloudcode": proxyEntry("https://cloudcode-pa.googleapis.com", "/api-proxy/google-cloudcode"),
      "/api-proxy/google-cloudcode-sandbox": proxyEntry("https://daily-cloudcode-pa.sandbox.googleapis.com", "/api-proxy/google-cloudcode-sandbox"),
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

      // pi-web-ui only exports "." + "./app.css". We deep-import from its dist
      // modules to avoid pulling the entire barrel (ChatPanel, artifacts, etc.).
      // This alias bypasses package.json "exports" restrictions.
      "@mariozechner/pi-web-ui/dist": path.resolve(__dirname, "node_modules/@mariozechner/pi-web-ui/dist"),
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
      // Externalize node:* imports (Rollup can't bundle them for the browser).
      // Note: do NOT externalize regular deps (e.g. @smithy/*). If they leak
      // through as bare imports, the built add-in will fail to boot.
      external: [
        /^node:/,
      ],
    },
  },
});
