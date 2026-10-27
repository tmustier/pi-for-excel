import { createHash, randomUUID } from "node:crypto";
import { execFileSync, spawnSync } from "node:child_process";
import { closeSync, lstatSync, mkdirSync, openSync, readdirSync, readFileSync, readlinkSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const args = new Set(process.argv.slice(2));
if ([...args].some(arg => !["--help", "--last", "--force"].includes(arg))) {
  console.error("Unknown option. Use --help.");
  process.exit(2);
}
// agent-default (not a user rule): 2026-10-27 — allow 10 minutes for the full gate.
if (args.has("--help")) {
  console.log(`Usage: npm run verify -- [--force | --last | --help]
Runs every npm run check sub-step, the full npm test suite, test:browser and build.
Failures do not stop later steps. Each failure includes an excerpt and full log path.
--last   Reprint the previous report and exit with its status, without running checks.
--force  Rerun even if this exact worktree content already passed.
Logs remain in this worktree's Git directory until the next run.
The pass cache covers tracked and non-ignored untracked file bytes and modes,
the commands and Node version. Git-ignored dependencies and environment changes
are not fingerprinted: use --force after changing them.
Give this command a timeout of at least 10 minutes.`);
  process.exit(0);
}

const git = (...command) => execFileSync("git", command, { cwd: root, encoding: "utf8" }).trimEnd();
const state = join(git("rev-parse", "--absolute-git-dir"), "verify");
const last = join(state, "last");
const reportPath = join(last, "report.json");

if (args.has("--last")) {
  try {
    const previous = JSON.parse(readFileSync(reportPath, "utf8"));
    console.log(previous.report);
    process.exit(previous.status);
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
    console.error("No previous verification report in this worktree.");
    process.exit(2);
  }
}

const { scripts } = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
// agent-default (not a user rule): 2026-10-27 — include Chromium on every full run.
const steps = scripts.check.split(" && ").map(command => {
  const match = /^npm run ([\w:-]+)$/.exec(command);
  if (!match) throw new Error(`Unsupported check sub-step: ${command}`);
  return match[1];
}).concat("test", "test:browser", "build");

function fingerprint() {
  const hash = createHash("sha256");
  hash.update(JSON.stringify([process.version, process.platform, process.arch, steps]));
  const files = new Set(git("ls-files", "-z", "--cached", "--others", "--exclude-standard").split("\0").filter(Boolean));
  for (const file of [...files].sort()) {
    const path = join(root, file);
    hash.update(`${file}\0`);
    let stat;
    try {
      stat = lstatSync(path);
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
      hash.update("deleted\0");
      continue;
    }
    hash.update(`${stat.mode}\0`);
    const bytes = stat.isSymbolicLink() ? Buffer.from(readlinkSync(path)) : readFileSync(path);
    hash.update(`${bytes.length}\0`);
    hash.update(bytes);
  }
  return hash.digest("hex");
}

mkdirSync(state, { recursive: true });
// Each run claims a file named after its process ID; claims left by killed runs are removed.
const owners = join(state, "owners");
mkdirSync(owners, { recursive: true });
const lock = join(owners, `${process.pid}.${randomUUID()}`);
writeFileSync(lock, "");
for (const owner of readdirSync(owners)) {
  const other = join(owners, owner);
  if (other === lock) continue;
  const pid = Number(owner.split(".")[0]);
  try {
    process.kill(pid, 0);
  } catch (error) {
    if (error.code === "ESRCH") {
      rmSync(other, { force: true });
      continue;
    }
    if (error.code !== "EPERM") throw error;
  }
  rmSync(lock, { force: true });
  console.error(`Verification is already running in this worktree (process ${pid}).`);
  process.exit(2);
}

try {
  const key = fingerprint();
  let previous;
  try {
    previous = JSON.parse(readFileSync(reportPath, "utf8"));
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
  if (!args.has("--force") && previous?.status === 0 && previous.key === key) {
    console.log("All checks already passed for this content; --force reruns, --last shows the report.");
  } else {
    rmSync(last, { recursive: true, force: true });
    mkdirSync(last);
    const results = [];
    for (const step of steps) {
      console.log(`Running ${step}…`);
      const start = performance.now();
      const log = join(last, `${step.replaceAll(":", "-")}.log`);
      const fd = openSync(log, "w");
      const result = spawnSync("npm", ["run", step], { cwd: root, stdio: ["ignore", fd, fd], shell: process.platform === "win32" });
      closeSync(fd);
      if (result.error) writeFileSync(log, String(result.error), { flag: "a" });
      results.push({ step, log, status: result.status ?? 1, seconds: ((performance.now() - start) / 1000).toFixed(1) });
    }
    const failed = results.filter(result => result.status !== 0);
    const lines = results.map(result => `${result.status === 0 ? "ok  " : "FAIL"}  ${result.step}  ${result.seconds}s`);
    // agent-default (not a user rule): 2026-10-27
    const excerptLines = 16;
    for (const result of failed) {
      const logLines = readFileSync(result.log, "utf8").split("\n");
      const firstFailure = logLines.findIndex(line => /^\s*(?:not ok\b|[✖×])|error TS\d|Error:/.test(line));
      const excerpt = firstFailure < 0 ? logLines.filter(line => line.trim()).slice(-excerptLines) : logLines.slice(firstFailure, firstFailure + excerptLines);
      lines.push(`\n--- ${result.step} (exit ${result.status}) ---`, ...excerpt, `Full log: ${result.log}`);
    }
    const status = failed.length ? 1 : 0;
    lines.push(failed.length ? `\nFailed: ${failed.map(result => result.step).join(", ")}` : "\nAll checks passed.");
    const report = lines.join("\n");
    writeFileSync(reportPath, JSON.stringify({ status, key: fingerprint() === key ? key : null, report }));
    console.log(report);
    process.exitCode = status;
  }
} finally {
  rmSync(lock, { force: true });
}
