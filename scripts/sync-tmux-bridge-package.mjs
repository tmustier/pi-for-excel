import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");

const sourceFiles = [
  "scripts/tmux-bridge-server.mjs",
];

const packageScriptsDir = path.join(repoRoot, "pkg", "tmux-bridge", "scripts");
fs.mkdirSync(packageScriptsDir, { recursive: true });

for (const relativePath of sourceFiles) {
  const sourcePath = path.join(repoRoot, relativePath);
  const destinationPath = path.join(packageScriptsDir, path.basename(relativePath));
  fs.copyFileSync(sourcePath, destinationPath);
  console.log(`[sync-tmux-bridge-package] copied ${relativePath} -> pkg/tmux-bridge/scripts/${path.basename(relativePath)}`);
}
