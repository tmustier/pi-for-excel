#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
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

function ensureMkcert() {
  if (commandExists("mkcert")) {
    return;
  }

  console.log("[pi-for-excel-proxy] mkcert not found.");

  if (process.platform === "darwin") {
    if (!commandExists("brew")) {
      console.error("[pi-for-excel-proxy] Homebrew is not installed.");
      console.error("[pi-for-excel-proxy] Install Homebrew first: https://brew.sh");
      process.exit(1);
    }

    console.log("[pi-for-excel-proxy] Installing mkcert via Homebrew...");
    run("brew", ["install", "mkcert"]);

    if (!commandExists("mkcert")) {
      console.error("[pi-for-excel-proxy] mkcert installation completed but mkcert is still unavailable in PATH.");
      process.exit(1);
    }

    return;
  }

  console.error("[pi-for-excel-proxy] Please install mkcert, then run this command again.");
  console.error("[pi-for-excel-proxy] Install instructions: https://github.com/FiloSottile/mkcert#installation");
  process.exit(1);
}

function ensureCertificates() {
  fs.mkdirSync(certDir, { recursive: true });

  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    return;
  }

  ensureMkcert();

  console.log("[pi-for-excel-proxy] Generating local HTTPS certificates...");
  run("mkcert", ["-install"]);
  run("mkcert", ["-key-file", keyPath, "-cert-file", certPath, "localhost"], {
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
  ensureCertificates();
}
startProxy(proxyConfig.proxyArgs);
