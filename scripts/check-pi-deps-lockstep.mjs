import { promises as fs } from "node:fs";

const PI_DEPENDENCIES = [
  "@mariozechner/pi-ai",
  "@mariozechner/pi-web-ui",
  "@mariozechner/pi-agent-core",
];

async function main() {
  const source = await fs.readFile("package.json", "utf8");
  const pkg = JSON.parse(source);
  const dependencies = pkg.dependencies ?? {};

  const entries = PI_DEPENDENCIES.map((name) => [name, dependencies[name]]);

  const missing = entries.filter(([, version]) => typeof version !== "string");
  if (missing.length > 0) {
    console.error("\n✗ Missing required Pi dependencies in package.json:\n");
    for (const [name] of missing) {
      console.error(`  - ${name}`);
    }
    process.exitCode = 1;
    return;
  }

  const versions = new Set(entries.map(([, version]) => version));
  if (versions.size > 1) {
    console.error("\n✗ Pi dependencies are out of lockstep in package.json:\n");
    for (const [name, version] of entries) {
      console.error(`  - ${name}: ${version}`);
    }
    console.error("\nExpected all three Pi package versions to match exactly.");
    process.exitCode = 1;
    return;
  }

  const lockstepVersion = entries[0]?.[1] ?? "(unknown)";
  console.log(`✓ Pi dependencies are in lockstep (${lockstepVersion}).`);
}

void main();
