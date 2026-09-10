import assert from "node:assert/strict";
import { test } from "node:test";

import { createAllTools } from "../src/tools/index.ts";
import { applyExperimentalToolGates } from "../src/tools/experimental-tool-gates.ts";
import { isTmuxBridgeDetails } from "../src/tools/tool-details.ts";

async function runtimeTools(dependencies: Parameters<typeof applyExperimentalToolGates>[1] = {}) {
  return applyExperimentalToolGates(createAllTools({ hostKind: "office" }), dependencies);
}

function findTool(tools: Awaited<ReturnType<typeof runtimeTools>>, name: string) {
  const tool = tools.find((candidate) => candidate.name === name);
  assert.ok(tool, `${name} should be visible to the model`);
  return tool;
}

void test("the model-visible tmux tool re-checks bridge availability for each execution", async () => {
  let bridgeUrl: string | undefined = "https://localhost:4441";
  const tools = await runtimeTools({
    getTmuxBridgeUrl: () => Promise.resolve(bridgeUrl),
    validateBridgeUrl: (url) => url,
    probeTmuxBridge: () => Promise.resolve(false),
  });
  const tmux = findTool(tools, "tmux");

  const unreachable = await tmux.execute("call-unreachable", { action: "list_sessions" });
  assert.ok(isTmuxBridgeDetails(unreachable.details));
  assert.equal(unreachable.details.gateReason, "bridge_unreachable");

  bridgeUrl = undefined;
  const missing = await tmux.execute("call-missing", { action: "list_sessions" });
  assert.ok(isTmuxBridgeDetails(missing.details));
  assert.equal(missing.details.gateReason, "missing_bridge_url");
});

void test("the model-visible files and Office.js tools remain available without feature flags", async () => {
  const tools = await runtimeTools({
    getExecutionMode: () => Promise.resolve("yolo"),
    requestOfficeJsExecuteApproval: () => Promise.reject(new Error("pure Office.js must not request approval in Auto mode")),
  });

  await assert.rejects(
    async () => findTool(tools, "files").execute("call-files", { action: "read" }),
    /'path' is required/u,
  );

  const officeResult = await findTool(tools, "execute_office_js").execute("call-office", {
    explanation: "Inspect workbook",
    code: "return Excel.run(async () => true);",
  });
  const officeText = officeResult.content[0]?.type === "text" ? officeResult.content[0].text : "";
  assert.match(officeText, /Do not call Excel\.run/u);
});
