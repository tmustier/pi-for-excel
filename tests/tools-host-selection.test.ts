import assert from "node:assert/strict";
import { afterEach, test } from "node:test";

import { createAllTools } from "../src/tools/index.ts";
import { UnsupportedHostToolError } from "../src/tools/unsupported-host-tool.ts";

interface WpsGlobals {
  Application?: DynamicValue;
}

const globals = globalThis as typeof globalThis & WpsGlobals;
const originalApplication = globals.Application;

afterEach(() => {
  if (originalApplication === undefined) delete globals.Application;
  else globals.Application = originalApplication;
});

function registeredTool(hostKind: "office" | "wps", name: string) {
  const tool = createAllTools({ hostKind }).find((candidate) => candidate.name === name);
  assert.ok(tool, `${name} should be registered for ${hostKind}`);
  return tool;
}

void test("registered Office.js tools fail publicly on WPS without entering Office.js", async () => {
  const wpsTool = registeredTool("wps", "execute_office_js");
  await assert.rejects(
    async () => wpsTool.execute("call-wps", {
      explanation: "Inspect workbook",
      code: "return { ok: true };",
    }),
    (error: DynamicValue) => {
      assert.ok(error instanceof UnsupportedHostToolError);
      assert.equal(error.hostKind, "wps");
      assert.equal(error.toolName, "execute_office_js");
      return true;
    },
  );

  const officeResult = await registeredTool("office", "execute_office_js").execute("call-office", {
    explanation: "Inspect workbook",
    code: "return Excel.run(async () => true);",
  });
  assert.equal(officeResult.content[0]?.type, "text");
  if (officeResult.content[0]?.type === "text") {
    assert.match(officeResult.content[0].text, /Do not call Excel\.run/u);
  }
});

void test("the registered WPS read tool executes against the WPS host while local skills stay usable", async () => {
  const range = {
    Address: "$A$1:$B$2",
    Value2: [["Region", "Sales"], ["North", 42]],
    Formula: [["", ""], ["", ""]],
    NumberFormat: [["General", "General"], ["General", "0"]],
    Rows: { Count: 2 },
    Columns: { Count: 2 },
  };
  const sheet = {
    Name: "Summary",
    Range: (address: string) => address === "A1:B2" ? range : null,
  };
  globals.Application = {
    ActiveWorkbook: { Name: "Plan.xlsx", Worksheets: { Count: 1, Item: () => sheet } },
    ActiveSheet: sheet,
  };

  const readResult = await registeredTool("wps", "read_range").execute("call-read", {
    range: "Summary!A1:B2",
  });
  const readText = readResult.content[0]?.type === "text" ? readResult.content[0].text : "";
  assert.match(readText, /Region/u);
  assert.match(readText, /North/u);
  assert.match(readText, /42/u);

  const skillsResult = await registeredTool("wps", "skills").execute("call-skills", { action: "list" });
  const skillsText = skillsResult.content[0]?.type === "text" ? skillsResult.content[0].text : "";
  assert.match(skillsText, /Available Agent Skills/u);
});
