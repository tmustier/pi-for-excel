import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentTool } from "@earendil-works/pi-agent-core";
import { Type } from "typebox";

import { withWorkbookCoordinator } from "../src/tools/with-workbook-coordinator.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";

const parameters = Type.Object({});

function writeTool(text: string, execute: () => Promise<void>): AgentTool {
  return {
    name: "write_cells",
    label: "Write cells",
    description: "write test boundary",
    parameters,
    execute: async () => {
      await execute();
      return { content: [{ type: "text", text }], details: {} };
    },
  };
}

function wrapped(tool: AgentTool, workbookId: string | null): AgentTool {
  const result = withWorkbookCoordinator(
    [tool],
    coordinator,
    { getWorkbookId: () => Promise.resolve(workbookId), getSessionId: () => "session" },
  )[0];
  if (!result) throw new Error("Expected wrapped tool");
  return result;
}

const coordinator = createWorkbookCoordinator();

function resultText(result: Awaited<ReturnType<AgentTool["execute"]>>): string {
  const block = result.content[0];
  return block?.type === "text" ? block.text : "";
}

void test("unknown-workbook writes return tool results serially under the shared lock", async () => {
  let releaseFirst: (() => void) | undefined;
  const firstBoundary = new Promise<void>((resolve) => { releaseFirst = resolve; });
  const first = wrapped(writeTool("first committed", () => firstBoundary), null);
  const second = wrapped(writeTool("second committed", () => Promise.resolve()), null);

  const firstResult = first.execute("first", {});
  const secondResult = second.execute("second", {});
  const whileLocked = await Promise.race([
    secondResult.then(() => "unexpected result"),
    new Promise<string>((resolve) => setTimeout(() => resolve("still locked"), 10)),
  ]);
  assert.equal(whileLocked, "still locked");

  releaseFirst?.();
  assert.equal(resultText(await firstResult), "first committed");
  assert.equal(resultText(await secondResult), "second committed");
});

void test("known workbook identities isolate tool results from an unrelated lock", async () => {
  let releaseFirst: (() => void) | undefined;
  const firstBoundary = new Promise<void>((resolve) => { releaseFirst = resolve; });
  const first = wrapped(writeTool("book A committed", () => firstBoundary), "book-a");
  const second = wrapped(writeTool("book B committed", () => Promise.resolve()), "book-b");

  const firstResult = first.execute("first", {});
  assert.equal(resultText(await second.execute("second", {})), "book B committed");
  releaseFirst?.();
  assert.equal(resultText(await firstResult), "book A committed");
});

void test("a failed unknown-workbook write releases the lock for the next tool result", async () => {
  const failure = wrapped(writeTool("unused", () => Promise.reject(new Error("host write failed"))), null);
  const next = wrapped(writeTool("next committed", () => Promise.resolve()), null);

  await assert.rejects(() => failure.execute("failure", {}), /host write failed/u);
  assert.equal(resultText(await next.execute("next", {})), "next committed");
});
