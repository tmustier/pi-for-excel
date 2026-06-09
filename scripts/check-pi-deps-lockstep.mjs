import { promises as fs } from "node:fs";

/**
 * Pi dependency policy:
 *
 * - `@earendil-works/pi-ai` and `@earendil-works/pi-agent-core` must stay in
 *   exact lockstep (same spec in package.json, same resolved version in
 *   package-lock.json).
 * - `@earendil-works/pi-web-ui` is allowed to lag: upstream stopped publishing
 *   it in lockstep after 0.75.3. It must still be exact-pinned.
 * - The lockfile must resolve exactly ONE copy of `pi-ai` and `pi-agent-core`
 *   (no nested duplicates). Duplicate pi-ai copies mean two model registries:
 *   the ModelSelector (pi-web-ui) and the app would disagree about available
 *   models. The root `overrides` entry for pi-ai keeps pi-web-ui's nested
 *   range deduped onto the root version — this check verifies the effect.
 */

const LOCKSTEP_PAIR = ["@earendil-works/pi-ai", "@earendil-works/pi-agent-core"];
const SINGLETON_PACKAGES = ["@earendil-works/pi-ai", "@earendil-works/pi-agent-core"];
const PI_DEPENDENCIES = [...LOCKSTEP_PAIR, "@earendil-works/pi-web-ui"];

const EXACT_VERSION_RE = /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/;

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

  console.error(`\n✗ Pi core dependencies are out of lockstep in ${sourceName}:\n`);
  for (const [name, version] of entries) {
    console.error(`  - ${name}: ${version}`);
  }
  console.error("\nExpected pi-ai and pi-agent-core versions to match exactly.");
  return true;
}

function failIfNotExactPins(entries) {
  const loose = entries.filter(([, version]) => !EXACT_VERSION_RE.test(version));
  if (loose.length === 0) return false;

  console.error("\n✗ Pi dependencies must be exact-pinned in package.json:\n");
  for (const [name, version] of loose) {
    console.error(`  - ${name}: ${version}`);
  }
  return true;
}

function failIfDuplicateResolutions(lockPackages) {
  let hasErrors = false;

  for (const name of SINGLETON_PACKAGES) {
    const suffix = `node_modules/${name}`;
    const resolutions = Object.keys(lockPackages).filter(
      (key) => key === suffix || key.endsWith(`/${suffix}`),
    );

    if (resolutions.length > 1) {
      console.error(`\n✗ ${name} resolves to multiple copies in package-lock.json:\n`);
      for (const key of resolutions) {
        console.error(`  - ${key}: ${lockPackages[key]?.version}`);
      }
      console.error(
        "\nExpected a single shared copy. Check the root \"overrides\" entry for pi-ai.",
      );
      hasErrors = true;
    }
  }

  return hasErrors;
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

  const pairFrom = (entries) => entries.filter(([name]) => LOCKSTEP_PAIR.includes(name));

  const hasErrors =
    failMissing("package.json", findMissing(packageJsonEntries)) ||
    failMissing("package-lock.json", findMissing(lockEntries)) ||
    failIfNotExactPins(packageJsonEntries) ||
    failIfNotLockstep("package.json", pairFrom(packageJsonEntries)) ||
    failIfNotLockstep("package-lock.json", pairFrom(lockEntries)) ||
    failIfDuplicateResolutions(lockPackages);

  if (hasErrors) {
    process.exitCode = 1;
    return;
  }

  const coreVersion = packageJsonEntries[0]?.[1] ?? "(unknown)";
  const webUiVersion =
    packageJsonEntries.find(([name]) => name === "@earendil-works/pi-web-ui")?.[1] ?? "(unknown)";
  console.log(
    `✓ Pi dependencies OK (pi-ai/pi-agent-core: ${coreVersion}, pi-web-ui: ${webUiVersion}, single shared pi-ai copy).`,
  );
}

void main();
