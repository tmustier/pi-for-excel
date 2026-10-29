import { defineConfig, type Plugin } from "vite";
import { isPiAuthRequestAllowed } from "./src/dev-auth-policy.js";
import { resolveDevOrigin } from "./scripts/generate-dev-manifest.mjs";
import type { IncomingMessage, ServerResponse } from "node:http";
import fs from "fs";
import path from "path";
import os from "os";

function piAuthPlugin(): Plugin {
  const authPath = path.join(os.homedir(), ".pi", "agent", "auth.json");
  const allowNonLocalHost = process.env.PI_AUTH_ALLOW_NONLOCAL_HOST === "1";

  return {
    name: "pi-auth",
    configureServer(server) {
      server.middlewares.use("/__pi-auth", (req: IncomingMessage, res: ServerResponse) => {
        // SECURITY: auth.json can contain API keys + refresh tokens.
        // Only serve it to loopback clients on loopback hostnames (Excel webviews,
        // local browser). QEMU user networking can make WPS guest traffic appear
        // loopback to Vite, so also require a localhost/127.0.0.1 Host header by default.
        const remote = req.socket?.remoteAddress;
        if (!isPiAuthRequestAllowed({ remoteAddress: remote, hostHeader: req.headers.host, authorityHeader: req.headers[":authority"], allowNonLocalHost })) {
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

type ProxyReqLike = {
  removeHeader(name: string): void;
  path?: string;
};
type ProxyServerLike = { on(event: "proxyReq", handler: (proxyReq: ProxyReqLike) => void): void };

function stripBrowserHeaders(proxy: ProxyServerLike) {
  proxy.on("proxyReq", (proxyReq) => {
    proxyReq.removeHeader("origin");
    proxyReq.removeHeader("referer");
    proxyReq.removeHeader("user-agent");
    proxyReq.removeHeader("accept-language");
    proxyReq.removeHeader("sec-fetch-mode");
    proxyReq.removeHeader("sec-fetch-site");
    proxyReq.removeHeader("sec-fetch-dest");
    proxyReq.removeHeader("sec-ch-ua");
    proxyReq.removeHeader("sec-ch-ua-mobile");
    proxyReq.removeHeader("sec-ch-ua-platform");
    proxyReq.removeHeader("anthropic-dangerous-direct-browser-access");

    if (typeof proxyReq.path === "string" && /%3a/i.test(proxyReq.path)) {
      proxyReq.path = proxyReq.path.replaceAll("%3A", ":").replaceAll("%3a", ":");
    }
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

function buildBrowserAliasMap(): Record<string, string> {
  const resolveFromRoot = (relativePath: string): string => path.resolve(__dirname, relativePath);

  return {
    // Stub Node.js built-ins imported by Anthropic SDK's transitive deps (undici, @smithy).
    // These code paths are never executed in the browser — all API calls use fetch().
    stream: resolveFromRoot("src/stubs/stream.ts"),

    ajv: resolveFromRoot("src/stubs/ajv.ts"),
    "ajv-formats": resolveFromRoot("src/stubs/ajv-formats.ts"),

  };
}

function buildBrowserAliases(): { find: string | RegExp; replacement: string }[] {
  return Object.entries(buildBrowserAliasMap()).map(([find, replacement]) => ({ find, replacement }));
}

interface DevProxyConfig {
  host: string;
  clientPort: number;
}

function resolveDevProxy(): DevProxyConfig | null {
  const devHost = process.env.DEV_HOST?.trim();
  const portlessUrl = process.env.PORTLESS_URL?.trim();
  if (!devHost && !portlessUrl) return null;

  const resolved = resolveDevOrigin({
    env: { DEV_HOST: devHost, PORTLESS_URL: portlessUrl },
  });
  const url = new URL(resolved.origin);
  return {
    host: url.hostname,
    clientPort: url.port ? Number.parseInt(url.port, 10) : 443,
  };
}

function parseDevServerPort(raw: string | undefined): number {
  const parsed = Number.parseInt(raw ?? "", 10);
  return Number.isInteger(parsed) && parsed > 0 && parsed < 65536 ? parsed : 3141;
}

const keyPath = path.resolve(__dirname, "key.pem");
const certPath = path.resolve(__dirname, "cert.pem");

const hasHttpsCerts = fs.existsSync(keyPath) && fs.existsSync(certPath);

const devProxy = resolveDevProxy();

export default defineConfig({
  plugins: [
    piAuthPlugin(),
  ],

  server: {
    ...(devProxy
      ? {
          host: "localhost",
          strictPort: true,
          port: parseDevServerPort(process.env.PORT),
          allowedHosts: [devProxy.host],
          hmr: {
            protocol: "wss",
            host: devProxy.host,
            clientPort: devProxy.clientPort,
          },
        }
      : {
          // Bind IPv6 too: Excel's webview may resolve localhost → ::1 and fail if we only listen on 127.0.0.1.
          host: "::",
          strictPort: true,
          port: 3141,
          https: hasHttpsCerts
            ? { key: fs.readFileSync(keyPath), cert: fs.readFileSync(certPath) }
            : undefined,
        }),

    proxy: {
      "/oauth-proxy/anthropic-platform": proxyEntry("https://platform.claude.com", "/oauth-proxy/anthropic-platform"),
      "/oauth-proxy/anthropic": proxyEntry("https://console.anthropic.com", "/oauth-proxy/anthropic"),
      "/oauth-proxy/github": proxyEntry("https://github.com", "/oauth-proxy/github"),

      "/api-proxy/github": proxyEntry("https://api.github.com", "/api-proxy/github"),
      "/api-proxy/anthropic": proxyEntry("https://api.anthropic.com", "/api-proxy/anthropic"),
      "/api-proxy/openai-auth": proxyEntry("https://auth.openai.com", "/api-proxy/openai-auth"),
      "/api-proxy/openai": proxyEntry("https://api.openai.com", "/api-proxy/openai"),
      "/api-proxy/chatgpt": proxyEntry("https://chatgpt.com", "/api-proxy/chatgpt"),
      "/api-proxy/google-oauth": proxyEntry("https://oauth2.googleapis.com", "/api-proxy/google-oauth"),
      "/api-proxy/google-cloudcode-sandbox": proxyEntry("https://daily-cloudcode-pa.sandbox.googleapis.com", "/api-proxy/google-cloudcode-sandbox"),
      "/api-proxy/google-cloudcode": proxyEntry("https://cloudcode-pa.googleapis.com", "/api-proxy/google-cloudcode"),
      "/api-proxy/google": proxyEntry("https://generativelanguage.googleapis.com", "/api-proxy/google"),

      "/model-catalog": {
        target: "https://pi.dev",
        changeOrigin: true,
        rewrite: (p: string) =>
          p.replace(/^\/model-catalog\/([0-9][0-9a-z.+-]*)\/([a-z0-9-]+)$/, "/api/models/providers/$2?types=chat&pi-version=$1"),
        secure: true,
        configure: stripBrowserHeaders,
      },
    },
  },

  define: {
    "process.env": "{}",
  },

  esbuild: { target: "esnext" },

  resolve: {
    alias: buildBrowserAliases(),
    dedupe: ["marked"],
  },

  build: {
    target: "esnext",
    commonjsOptions: {
      ignoreDynamicRequires: true,
    },
    rollupOptions: {
      input: {
        taskpane: "src/taskpane.html",
        "ui-gallery": "src/ui-gallery.html",
      },
      external: [
        /^node:/,
      ],
    },
  },
});
