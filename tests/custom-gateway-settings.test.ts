import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";

void test("custom gateway delete uses the app overlay confirmation instead of native confirm", async () => {
  const source = await readFile(
    new URL("../src/commands/builtins/custom-gateway-settings.ts", import.meta.url),
    "utf8",
  );

  assert.match(source, /requestConfirmationDialog/);
  assert.match(source, /confirmButtonTone:\s*"danger"/);
  assert.match(source, /restoreFocusOnClose:\s*false/);
  assert.match(source, /deleteOpenAiGatewayConfig/);
  assert.doesNotMatch(source, /\bconfirm\s*\(/);
});
