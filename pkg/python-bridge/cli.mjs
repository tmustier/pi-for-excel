#!/usr/bin/env node

import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const PACKAGE_TAG = "pi-for-excel-python-bridge";
const DEFAULT_PORT = "3340";
const INSTALL_MISSING_FLAG = "--install-missing";

const cliDir = path.dirname(fileURLToPath(import.meta.url));
const bridgeScriptPath = path.join(cliDir, "scripts", "python-bridge-server.mjs");

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

function canRunBinary(command, args = ["--version"]) {
  const result = spawnSync(command, args, {
    stdio: "ignore",
  });

  if (result.error || result.signal) {
    return false;
  }

  return result.status === 0;
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    stdio: "inherit",
    ...options,
  });

  if (result.error) {
    console.error(`[${PACKAGE_TAG}] Failed to run: ${command}`);
    console.error(result.error.message);
    process.exit(1);
  }

  if (typeof result.status === "number" && result.status !== 0) {
    process.exit(result.status);
  }

  if (result.signal) {
    console.error(`[${PACKAGE_TAG}] ${command} terminated by signal ${result.signal}`);
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
      console.error(`[${PACKAGE_TAG}] Homebrew is not installed.`);
      console.error(`[${PACKAGE_TAG}] Install Homebrew first: https://brew.sh`);
      process.exit(1);
    }

    console.log(`[${PACKAGE_TAG}] Installing mkcert via Homebrew...`);
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

    console.error(`[${PACKAGE_TAG}] mkcert is installed but not compatible with required CLI flags.`);
    console.error(`[${PACKAGE_TAG}] Ensure FiloSottile mkcert is used (not the npm mkcert package).`);
    process.exit(1);
  }

  console.error(`[${PACKAGE_TAG}] Please install mkcert, then run this command again.`);
  console.error(`[${PACKAGE_TAG}] Install instructions: https://github.com/FiloSottile/mkcert#installation`);
  process.exit(1);
}

function installMkcertCa(mkcertCommand) {
  const result = spawnSync(mkcertCommand, ["-install"], {
    stdio: "inherit",
  });

  if (!result.error && result.status === 0 && !result.signal) {
    return;
  }

  console.error(`[${PACKAGE_TAG}] Failed to install mkcert local CA.`);
  console.error(`[${PACKAGE_TAG}] Run manually: mkcert -install`);
  console.error(`[${PACKAGE_TAG}] If it fails, fix trust-store permissions and retry.`);

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

  console.log(`[${PACKAGE_TAG}] Generating local HTTPS certificates...`);
  installMkcertCa(mkcertCommand);

  run(mkcertCommand, ["-key-file", keyPath, "-cert-file", certPath, "localhost"], {
    cwd: certDir,
  });

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error(`[${PACKAGE_TAG}] Failed to generate TLS certificates.`);
    process.exit(1);
  }
}

function resolveConfiguredLibreOfficeBinary(env = process.env) {
  const raw = typeof env.PYTHON_BRIDGE_LIBREOFFICE_BIN === "string"
    ? env.PYTHON_BRIDGE_LIBREOFFICE_BIN
    : "";
  const trimmed = raw.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function resolveBundledLibreOfficeBinary() {
  const candidates = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    path.join(homeDir, "Applications", "LibreOffice.app", "Contents", "MacOS", "soffice"),
  ];

  for (const candidate of candidates) {
    if (!fs.existsSync(candidate)) {
      continue;
    }

    if (canRunBinary(candidate, ["--version"])) {
      return candidate;
    }
  }

  return null;
}

function resolveAvailableLibreOfficeBinary(env = process.env) {
  const configured = resolveConfiguredLibreOfficeBinary(env);
  if (configured && canRunBinary(configured, ["--version"])) {
    return configured;
  }

  if (canRunBinary("soffice", ["--version"])) {
    return "soffice";
  }

  if (canRunBinary("libreoffice", ["--version"])) {
    return "libreoffice";
  }

  return resolveBundledLibreOfficeBinary();
}

function applyLibreOfficeBinaryOverride(env) {
  if (resolveConfiguredLibreOfficeBinary(env)) {
    return;
  }

  const availableBinary = resolveAvailableLibreOfficeBinary(env);
  if (!availableBinary) {
    return;
  }

  if (availableBinary === "soffice" || availableBinary === "libreoffice") {
    return;
  }

  env.PYTHON_BRIDGE_LIBREOFFICE_BIN = availableBinary;
  console.log(`[${PACKAGE_TAG}] Using LibreOffice binary: ${availableBinary}`);
}

function installMissingDependencies() {
  const pythonMissing = !canRunBinary("python3", ["--version"]);
  const libreOfficeMissing = !resolveAvailableLibreOfficeBinary();

  if (!pythonMissing && !libreOfficeMissing) {
    return;
  }

  if (process.platform !== "darwin") {
    console.warn(`[${PACKAGE_TAG}] ${INSTALL_MISSING_FLAG} currently supports macOS/Homebrew only.`);
    if (pythonMissing) {
      console.warn(`[${PACKAGE_TAG}] Please install python3 manually and retry.`);
    }
    if (libreOfficeMissing) {
      console.warn(`[${PACKAGE_TAG}] Please install LibreOffice manually and retry.`);
    }
    return;
  }

  if (!commandExists("brew")) {
    console.error(`[${PACKAGE_TAG}] Homebrew is required for ${INSTALL_MISSING_FLAG}.`);
    console.error(`[${PACKAGE_TAG}] Install Homebrew first: https://brew.sh`);
    process.exit(1);
  }

  if (pythonMissing) {
    console.log(`[${PACKAGE_TAG}] Installing missing dependency: python3`);
    run("brew", ["install", "python"]);
  }

  if (libreOfficeMissing) {
    console.log(`[${PACKAGE_TAG}] Installing missing dependency: LibreOffice`);
    run("brew", ["install", "--cask", "libreoffice"]);
  }

  if (!canRunBinary("python3", ["--version"])) {
    console.warn(`[${PACKAGE_TAG}] python3 is still unavailable after install attempt.`);
  }

  if (!resolveAvailableLibreOfficeBinary()) {
    console.warn(`[${PACKAGE_TAG}] LibreOffice is still unavailable after install attempt.`);
    console.warn(`[${PACKAGE_TAG}] You can set PYTHON_BRIDGE_LIBREOFFICE_BIN to an absolute soffice path.`);
  }
}

function resolveBridgeConfig() {
  const userArgs = process.argv.slice(2);
  const installMissing = userArgs.includes(INSTALL_MISSING_FLAG);
  const bridgeUserArgs = userArgs.filter((arg) => arg !== INSTALL_MISSING_FLAG);

  const hasExplicitScheme = bridgeUserArgs.includes("--https") || bridgeUserArgs.includes("--http");
  const bridgeArgs = hasExplicitScheme ? bridgeUserArgs : ["--https", ...bridgeUserArgs];

  const usesHttpOnly = bridgeArgs.includes("--http") && !bridgeArgs.includes("--https");

  return {
    bridgeArgs,
    usesHttps: !usesHttpOnly,
    installMissing,
  };
}

function applyDefaultPort(env) {
  const configuredPort = typeof env.PORT === "string" ? env.PORT.trim() : "";
  if (configuredPort.length === 0) {
    env.PORT = DEFAULT_PORT;
  }
}

function startBridge(bridgeArgs) {
  fs.mkdirSync(certDir, { recursive: true });
  console.log(`[${PACKAGE_TAG}] Using certificate directory: ${certDir}`);

  const childEnv = { ...process.env };
  applyDefaultPort(childEnv);
  applyLibreOfficeBinaryOverride(childEnv);

  if (typeof childEnv.PI_FOR_EXCEL_CERT_DIR !== "string" || childEnv.PI_FOR_EXCEL_CERT_DIR.trim().length === 0) {
    childEnv.PI_FOR_EXCEL_CERT_DIR = certDir;
  }

  const child = spawn(process.execPath, [bridgeScriptPath, ...bridgeArgs], {
    env: childEnv,
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
    console.error(`[${PACKAGE_TAG}] Failed to start bridge process.`);
    console.error(error.message);
    process.exit(1);
  });
}

if (!fs.existsSync(bridgeScriptPath)) {
  console.error(`[${PACKAGE_TAG}] Missing bridge runtime files.`);
  console.error(`[${PACKAGE_TAG}] Reinstall the package or run npm pack again.`);
  process.exit(1);
}

const bridgeConfig = resolveBridgeConfig();
if (bridgeConfig.installMissing) {
  installMissingDependencies();
}
if (bridgeConfig.usesHttps) {
  ensureCertificates();
}
startBridge(bridgeConfig.bridgeArgs);
