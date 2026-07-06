import { execSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

function exec(cmd) {
  return execSync(cmd, { stdio: ["ignore", "pipe", "pipe"] }).toString().trim();
}

function tryExec(cmd) {
  try {
    return exec(cmd);
  } catch {
    return null;
  }
}

// Not running in a git checkout (e.g. npm pack / CI artifact) → nothing to do.
if (!tryExec("git --version") || !tryExec("git rev-parse --is-inside-work-tree")) {
  process.exit(0);
}

const hooksPath = tryExec("git config --get core.hooksPath") ?? "";

if (!hooksPath) {
  exec("git config core.hooksPath .githooks");
  // eslint-disable-next-line no-console -- setup script must report hook installation status
  console.log("Configured git core.hooksPath = .githooks");
} else if (hooksPath !== ".githooks") {
  // eslint-disable-next-line no-console -- setup script must report existing hook configuration
  console.log(
    `git core.hooksPath already set to \"${hooksPath}\"; leaving unchanged (repo expects .githooks).`,
  );
}

// Best-effort: ensure hooks are executable (POSIX).
try {
  const hookNames = ["pre-commit", "pre-push"];
  for (const hookName of hookNames) {
    const hookFile = path.resolve(".githooks", hookName);
    if (fs.existsSync(hookFile)) {
      fs.chmodSync(hookFile, 0o755);
    }
  }
} catch {
  // Ignore (e.g. Windows filesystems)
}
