import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { resolveVercelIgnoreCommandExitCode } from "../scripts/vercel-ignore-command.mjs";

const EXPECTED_IGNORE_COMMAND = "node scripts/vercel-ignore-command.mjs";

function isVercelIgnoreCommandTestPayloadShape(value) {
  return typeof value === "object" && value !== null;
}

async function readIgnoreCommand() {
  const raw = await readFile(new URL("../vercel.json", import.meta.url), "utf8");
  const parsed = JSON.parse(raw);

  if (!isVercelIgnoreCommandTestPayloadShape(parsed) || typeof parsed.ignoreCommand !== "string") {
    throw new Error("vercel.json is missing a string ignoreCommand");
  }

  return parsed.ignoreCommand;
}

test("vercel.json wires ignoreCommand to the deploy policy script", async () => {
  const ignoreCommand = await readIgnoreCommand();
  assert.equal(ignoreCommand, EXPECTED_IGNORE_COMMAND);
});

test("ignoreCommand allows manual deploys", () => {
  const exitCode = resolveVercelIgnoreCommandExitCode({});
  assert.equal(exitCode, 1, "manual deploys should build");
});

test("ignoreCommand allows main deploys", () => {
  const exitCode = resolveVercelIgnoreCommandExitCode({
    VERCEL_GIT_COMMIT_REF: "main",
  });

  assert.equal(exitCode, 1, "main branch deploys should build");
});

test("ignoreCommand allows pull request deploys", () => {
  const exitCode = resolveVercelIgnoreCommandExitCode({
    VERCEL_GIT_COMMIT_REF: "feature/re-enable-auto-deploy",
    VERCEL_GIT_PULL_REQUEST_ID: "290",
  });

  assert.equal(exitCode, 1, "pull request deploys should build");
});

test("ignoreCommand skips non-PR feature branches", () => {
  const exitCode = resolveVercelIgnoreCommandExitCode({
    VERCEL_GIT_COMMIT_REF: "feature/re-enable-auto-deploy",
  });

  assert.equal(exitCode, 0, "non-PR feature branch deploys should be skipped");
});
