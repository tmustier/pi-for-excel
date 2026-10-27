// Static contract: vercel.json must remain wired to the deploy-policy script;
// script behavior alone cannot detect a missing or redirected configuration dependency.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

import { resolveVercelIgnoreCommandExitCode } from "../scripts/vercel-ignore-command.mjs";

const EXPECTED_IGNORE_COMMAND = "node scripts/vercel-ignore-command.mjs";

function isVercelIgnoreCommandTestPayloadShape(value) {
  return typeof value === "object" && value !== null;
}

async function readVercelConfig() {
  const raw = await readFile(new URL("../vercel.json", import.meta.url), "utf8");
  const parsed = JSON.parse(raw);

  if (!isVercelIgnoreCommandTestPayloadShape(parsed) || typeof parsed.ignoreCommand !== "string") {
    throw new Error("vercel.json is missing a string ignoreCommand");
  }

  return parsed;
}

test("vercel.json wires ignoreCommand to the deploy policy script", async () => {
  const config = await readVercelConfig();
  assert.equal(config.ignoreCommand, EXPECTED_IGNORE_COMMAND);
});

test("only specification and non-PR integration branches disable automatic deployments", async () => {
  const config = await readVercelConfig();
  // Unspecified branches stay enabled in Vercel: main, human PRs (including
  // intermediate stack layers), and dependency PRs retain their previews.
  assert.deepEqual(config.git, {
    deploymentEnabled: {
      "product-description": false,
      "durable-plan": false,
      "integration-fixes": false,
    },
  });
});

test("deployment policy maps each Vercel context to deploy or skip", () => {
  const cases = [
    { name: "manual deploy", env: {}, expected: 1 },
    { name: "main branch", env: { VERCEL_GIT_COMMIT_REF: "main" }, expected: 1 },
    {
      name: "pull request",
      env: {
        VERCEL_GIT_COMMIT_REF: "feature/re-enable-auto-deploy",
        VERCEL_GIT_PULL_REQUEST_ID: "290",
      },
      expected: 1,
    },
    {
      name: "non-PR feature branch",
      env: { VERCEL_GIT_COMMIT_REF: "feature/re-enable-auto-deploy" },
      expected: 0,
    },
  ];

  for (const { name, env, expected } of cases) {
    assert.equal(resolveVercelIgnoreCommandExitCode(env), expected, name);
  }
});
