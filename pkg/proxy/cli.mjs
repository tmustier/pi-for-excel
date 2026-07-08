#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const cliDir = path.dirname(fileURLToPath(import.meta.url));
const proxyScriptPath = path.join(cliDir, "scripts", "cors-proxy-server.mjs");

const homeDir = os.homedir();
const appDir = path.join(homeDir, ".pi-for-excel");
const certDir = path.join(appDir, "certs");
const keyPath = path.join(certDir, "key.pem");
const certPath = path.join(certDir, "cert.pem");
const DEFAULT_PROXY_PORT = "3003";
const DEFAULT_PROXY_URL = `https://localhost:${DEFAULT_PROXY_PORT}`;

function commandExists(command) {
  const whichCommand = process.platform === "win32" ? "where" : "which";
  const result = spawnSync(whichCommand, [command], { stdio: "ignore" });
  return result.status === 0;
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    stdio: "inherit",
    ...options,
  });

  if (result.error) {
    console.error(`[pi-for-excel-proxy] Failed to run: ${command}`);
    console.error(result.error.message);
    process.exit(1);
  }

  if (typeof result.status === "number" && result.status !== 0) {
    process.exit(result.status);
  }

  if (result.signal) {
    console.error(`[pi-for-excel-proxy] ${command} terminated by signal ${result.signal}`);
    process.exit(1);
  }
}

function supportsMkcertCli(command) {
  const result = spawnSync(command, ["-CAROOT"], {
    stdio: "ignore",
  });

  if (result.error) {
    return false;
  }

  return result.status === 0 && !result.signal;
}

function resolveMkcertCommand() {
  const candidates = [];

  if (process.platform === "darwin") {
    const brewCandidates = ["/opt/homebrew/bin/mkcert", "/usr/local/bin/mkcert"];
    for (const candidate of brewCandidates) {
      if (fs.existsSync(candidate)) {
        candidates.push(candidate);
      }
    }
  }

  if (commandExists("mkcert")) {
    candidates.push("mkcert");
  }

  for (const candidate of candidates) {
    if (supportsMkcertCli(candidate)) {
      return candidate;
    }
  }

  if (process.platform === "darwin") {
    if (!commandExists("brew")) {
      console.error("[pi-for-excel-proxy] Homebrew is not installed.");
      console.error("[pi-for-excel-proxy] Install Homebrew first: https://brew.sh");
      process.exit(1);
    }

    console.log("[pi-for-excel-proxy] Installing mkcert via Homebrew...");
    run("brew", ["install", "mkcert"]);

    const brewCandidates = ["/opt/homebrew/bin/mkcert", "/usr/local/bin/mkcert", "mkcert"];
    for (const candidate of brewCandidates) {
      if (candidate !== "mkcert" && !fs.existsSync(candidate)) {
        continue;
      }

      if (supportsMkcertCli(candidate)) {
        return candidate;
      }
    }

    console.error("[pi-for-excel-proxy] mkcert is installed but not compatible with required CLI flags.");
    console.error("[pi-for-excel-proxy] Ensure FiloSottile mkcert is used (not the npm mkcert package).");
    process.exit(1);
  }

  console.error("[pi-for-excel-proxy] Please install mkcert, then run this command again.");
  console.error("[pi-for-excel-proxy] Install instructions: https://github.com/FiloSottile/mkcert#installation");
  process.exit(1);
}

function installMkcertCa(mkcertCommand) {
  const result = spawnSync(mkcertCommand, ["-install"], {
    stdio: "inherit",
  });

  if (!result.error && result.status === 0 && !result.signal) {
    return;
  }

  console.error("[pi-for-excel-proxy] Failed to install mkcert local CA.");
  console.error("[pi-for-excel-proxy] Run manually: mkcert -install");
  console.error("[pi-for-excel-proxy] If it fails, fix trust-store permissions and retry.");

  if (typeof result.status === "number" && result.status !== 0) {
    process.exit(result.status);
  }

  process.exit(1);
}

function ensureCertificates() {
  fs.mkdirSync(certDir, { recursive: true });

  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    return;
  }

  const mkcertCommand = resolveMkcertCommand();

  console.log("[pi-for-excel-proxy] Generating local HTTPS certificates...");
  installMkcertCa(mkcertCommand);

  run(mkcertCommand, ["-key-file", keyPath, "-cert-file", certPath, "localhost"], {
    cwd: certDir,
  });

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error("[pi-for-excel-proxy] Failed to generate TLS certificates.");
    process.exit(1);
  }
}

function resolveProxyConfig() {
  const userArgs = process.argv.slice(2);
  const hasExplicitScheme = userArgs.includes("--https") || userArgs.includes("--http");
  const proxyArgs = hasExplicitScheme ? userArgs : ["--https", ...userArgs];

  const usesHttpOnly = proxyArgs.includes("--http") && !proxyArgs.includes("--https");

  return {
    proxyArgs,
    usesHttps: !usesHttpOnly,
  };
}

function hasExplicitPort() {
  return typeof process.env.PORT === "string" && process.env.PORT.trim().length > 0;
}

function probeHttpsHealth(urlString) {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (healthy) => {
      if (settled) return;
      settled = true;
      resolve(healthy);
    };

    const req = https.request(
      urlString,
      {
        method: "GET",
        rejectUnauthorized: false,
        timeout: 800,
      },
      (res) => {
        let body = "";
        res.setEncoding("utf8");
        res.on("data", (chunk) => {
          body += chunk;
          if (body.length > 32) {
            req.destroy();
            finish(false);
          }
        });
        res.on("end", () => {
          finish(res.statusCode === 200 && body.trim() === "ok");
        });
      },
    );

    req.on("timeout", () => {
      req.destroy();
      finish(false);
    });
    req.on("error", () => finish(false));
    req.end();
  });
}

async function exitIfDefaultProxyAlreadyRunning(proxyConfig) {
  if (!proxyConfig.usesHttps || hasExplicitPort()) {
    return;
  }

  const [localhostHealthy, ipv4Healthy] = await Promise.all([
    probeHttpsHealth(`${DEFAULT_PROXY_URL}/healthz`),
    probeHttpsHealth(`https://127.0.0.1:${DEFAULT_PROXY_PORT}/healthz`),
  ]);

  if (localhostHealthy && ipv4Healthy) {
    console.log(`[pi-for-excel-proxy] Proxy already running at ${DEFAULT_PROXY_URL}`);
    console.log("[pi-for-excel-proxy] Nothing else to start — keep the existing proxy terminal open.");
    process.exit(0);
  }

  if (localhostHealthy || ipv4Healthy) {
    console.error("[pi-for-excel-proxy] Port 3003 has a partial/stale proxy listener.");
    console.error(`[pi-for-excel-proxy] ${DEFAULT_PROXY_URL}/healthz: ${localhostHealthy ? "ok" : "failed"}`);
    console.error(`[pi-for-excel-proxy] https://127.0.0.1:${DEFAULT_PROXY_PORT}/healthz: ${ipv4Healthy ? "ok" : "failed"}`);
    console.error("[pi-for-excel-proxy] Stop old pi-for-excel proxy processes, then run npx pi-for-excel-proxy again.");
    console.error("[pi-for-excel-proxy] Or set PORT=<free-port> and copy that URL into Pi for Excel /settings → Proxy.");
    process.exit(1);
  }
}

function startProxy(proxyArgs) {
  fs.mkdirSync(certDir, { recursive: true });
  console.log(`[pi-for-excel-proxy] Using certificate directory: ${certDir}`);

  const child = spawn(process.execPath, [proxyScriptPath, ...proxyArgs], {
    cwd: certDir,
    env: process.env,
    stdio: "inherit",
  });

  let shuttingDown = false;

  const forwardSignal = (signal) => {
    if (shuttingDown) {
      return;
    }
    shuttingDown = true;
    if (!child.killed) {
      child.kill(signal);
    }
  };

  process.on("SIGINT", () => forwardSignal("SIGINT"));
  process.on("SIGTERM", () => forwardSignal("SIGTERM"));

  child.on("exit", (code, signal) => {
    if (signal) {
      process.kill(process.pid, signal);
      return;
    }
    process.exit(code ?? 1);
  });

  child.on("error", (error) => {
    console.error("[pi-for-excel-proxy] Failed to start proxy process.");
    console.error(error.message);
    process.exit(1);
  });
}

if (!fs.existsSync(proxyScriptPath)) {
  console.error("[pi-for-excel-proxy] Missing proxy runtime files.");
  console.error("[pi-for-excel-proxy] Reinstall the package or run npm pack again.");
  process.exit(1);
}

const proxyConfig = resolveProxyConfig();
if (proxyConfig.usesHttps) {
  await exitIfDefaultProxyAlreadyRunning(proxyConfig);
  ensureCertificates();
}
startProxy(proxyConfig.proxyArgs);
