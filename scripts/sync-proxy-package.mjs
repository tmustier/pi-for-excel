#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");

const sourceFiles = [
  "scripts/cors-proxy-server.mjs",
  "scripts/proxy-target-policy.mjs",
];

const packageScriptsDir = path.join(repoRoot, "pkg", "proxy", "scripts");
fs.mkdirSync(packageScriptsDir, { recursive: true });

for (const relativePath of sourceFiles) {
  const sourcePath = path.join(repoRoot, relativePath);
  const destinationPath = path.join(packageScriptsDir, path.basename(relativePath));
  fs.copyFileSync(sourcePath, destinationPath);
  console.log(`[sync-proxy-package] copied ${relativePath} -> pkg/proxy/scripts/${path.basename(relativePath)}`);
}
