import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";

function fixture(t) {
  const root = realpathSync(mkdtempSync(join(tmpdir(), "pi-verify-")));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  mkdirSync(join(root, "scripts"));
  copyFileSync(new URL("../scripts/verify.mjs", import.meta.url), join(root, "scripts/verify.mjs"));
  const put = (file, content) => writeFileSync(join(root, file), content);
  const git = (...args) => execFileSync("git", args, { cwd: root, stdio: "pipe" });
  put("package.json", JSON.stringify({ scripts: {
    verify: "node scripts/verify.mjs",
    check: "npm run lint && npm run typecheck",
    lint: "node step.mjs lint",
    typecheck: "node step.mjs typecheck",
    test: "node step.mjs test",
    "test:browser": "node step.mjs test:browser",
    build: "node step.mjs build",
  } }));
  put("step.mjs", `import { appendFileSync, readFileSync } from "node:fs";
const step = process.argv[2];
appendFileSync(".git/calls", step + "\\n");
console.log("✔ a passing title contains ×");
console.log("setup noise\\n".repeat(20));
if (readFileSync("failures", "utf8").split("\\n").includes(step)) {
  console.log("✖ output from " + step);
  console.log("trailing noise\\n".repeat(20));
  process.exit(17);
}
`);
  put("failures", "");
  put("tracked", "before");
  git("init", "-q");
  git("add", ".");
  git("-c", "user.name=Test", "-c", "user.email=test@example.invalid", "-c", "commit.gpgsign=false", "commit", "-qm", "fixture");
  const run = (...args) => spawnSync("npm", ["run", "--silent", "verify", "--", ...args], { cwd: root, encoding: "utf8", shell: process.platform === "win32" });
  const calls = () => readFileSync(join(root, ".git/calls"), "utf8");
  return { root, put, run, calls };
}

test("verification reports every step and propagates failures without stopping", t => {
  const { root, put, run, calls } = fixture(t);
  put("failures", "lint\ntest");
  const result = run();
  assert.equal(result.status, 1, result.stdout + result.stderr);
  assert.equal(calls(), "lint\ntypecheck\ntest\ntest:browser\nbuild\n");
  for (const step of ["lint", "typecheck", "test", "test:browser", "build"]) {
    assert.match(result.stdout, new RegExp(`^${["lint", "test"].includes(step) ? "FAIL" : "ok  "}  ${step}  `, "m"));
  }
  for (const step of ["lint", "test"]) {
    assert.ok(result.stdout.includes(`output from ${step}`));
    const log = join(root, ".git/verify/last", `${step}.log`);
    assert.ok(result.stdout.includes(`Full log: ${log}`));
    assert.match(readFileSync(log, "utf8"), /output from/);
  }
  assert.equal(run().status, 1, "a failed run must not be cached");
  assert.equal(calls().split("\n").filter(Boolean).length, 10);
});

test("--last reprints the previous failure and status without running steps", t => {
  const { put, run, calls } = fixture(t);
  assert.equal(run("--last").status, 2);
  put("failures", "typecheck");
  const first = run();
  const before = calls();
  put("package.json", JSON.stringify({ scripts: { verify: "node scripts/verify.mjs", check: "unsupported command" } }));
  const last = run("--last");
  assert.equal(last.status, 1, last.stdout + last.stderr);
  assert.equal(last.stdout, first.stdout.split("\n").filter(line => !line.startsWith("Running ")).join("\n"));
  assert.equal(calls(), before);
});

test("identical content reuses a pass and --force reruns with logs preserved", t => {
  const { root, run, calls } = fixture(t);
  assert.equal(run().status, 0);
  const before = calls();
  const log = readFileSync(join(root, ".git/verify/last/build.log"), "utf8");
  const cached = run();
  assert.equal(cached.status, 0, cached.stdout + cached.stderr);
  assert.match(cached.stdout, /already passed/);
  assert.equal(calls(), before);
  assert.equal(readFileSync(join(root, ".git/verify/last/build.log"), "utf8"), log);
  assert.equal(run("--force").status, 0);
  assert.equal(calls(), before + before);
  assert.equal(run("--last").status, 0);
});

test("tracked edits, untracked files and deletions invalidate a cached pass", t => {
  const { root, put, run, calls } = fixture(t);
  assert.equal(run().status, 0);
  let count = 5;
  for (const change of [
    () => put("tracked", "edited"),
    () => put("new file", "untracked"),
    () => put("new file", "new bytes"),
    () => rmSync(join(root, "tracked")),
  ]) {
    change();
    const result = run();
    assert.equal(result.status, 0, result.stdout + result.stderr);
    assert.doesNotMatch(result.stdout, /already passed/);
    count += 5;
    assert.equal(calls().split("\n").filter(Boolean).length, count);
  }
});
