import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");

const sourceFiles = [
  "scripts/python-bridge-server.mjs",
];

const packageScriptsDir = path.join(repoRoot, "pkg", "python-bridge", "scripts");
fs.mkdirSync(packageScriptsDir, { recursive: true });

for (const relativePath of sourceFiles) {
  const sourcePath = path.join(repoRoot, relativePath);
  const destinationPath = path.join(packageScriptsDir, path.basename(relativePath));
  fs.copyFileSync(sourcePath, destinationPath);
  console.log(`[sync-python-bridge-package] copied ${relativePath} -> pkg/python-bridge/scripts/${path.basename(relativePath)}`);
}
