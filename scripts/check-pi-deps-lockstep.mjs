import { promises as fs } from "node:fs";

const PI_DEPENDENCIES = [
  "@mariozechner/pi-ai",
  "@mariozechner/pi-web-ui",
  "@mariozechner/pi-agent-core",
];

function findMissing(entries) {
  return entries.filter(([, version]) => typeof version !== "string");
}

function failMissing(fileName, missingEntries) {
  if (missingEntries.length === 0) return false;

  console.error(`\n✗ Missing required Pi dependencies in ${fileName}:\n`);
  for (const [name] of missingEntries) {
    console.error(`  - ${name}`);
  }

  return true;
}

function failIfNotLockstep(sourceName, entries) {
  const versions = new Set(entries.map(([, version]) => version));
  if (versions.size <= 1) return false;

  console.error(`\n✗ Pi dependencies are out of lockstep in ${sourceName}:\n`);
  for (const [name, version] of entries) {
    console.error(`  - ${name}: ${version}`);
  }
  console.error("\nExpected all three Pi package versions to match exactly.");
  return true;
}

async function main() {
  const [packageJsonSource, packageLockSource] = await Promise.all([
    fs.readFile("package.json", "utf8"),
    fs.readFile("package-lock.json", "utf8"),
  ]);

  const pkg = JSON.parse(packageJsonSource);
  const lock = JSON.parse(packageLockSource);

  const packageJsonDependencies = pkg.dependencies ?? {};
  const packageJsonEntries = PI_DEPENDENCIES.map((name) => [name, packageJsonDependencies[name]]);

  const lockPackages = lock.packages ?? {};
  const lockEntries = PI_DEPENDENCIES.map((name) => [
    name,
    lockPackages[`node_modules/${name}`]?.version,
  ]);

  const hasErrors =
    failMissing("package.json", findMissing(packageJsonEntries)) ||
    failMissing("package-lock.json", findMissing(lockEntries)) ||
    failIfNotLockstep("package.json", packageJsonEntries) ||
    failIfNotLockstep("package-lock.json", lockEntries);

  if (hasErrors) {
    process.exitCode = 1;
    return;
  }

  const specVersion = packageJsonEntries[0]?.[1] ?? "(unknown)";
  const resolvedVersion = lockEntries[0]?.[1] ?? "(unknown)";
  console.log(`✓ Pi dependencies are in lockstep (spec: ${specVersion}, resolved: ${resolvedVersion}).`);
}

void main();
