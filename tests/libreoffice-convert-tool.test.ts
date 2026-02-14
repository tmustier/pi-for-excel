import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentToolResult } from "@mariozechner/pi-agent-core";

import {
  createLibreOfficeConvertTool,
  type LibreOfficeBridgeConfig,
  type LibreOfficeConvertRequest,
  type LibreOfficeConvertResponse,
  type LibreOfficeConvertToolDetails,
} from "../src/tools/libreoffice-convert.ts";

function firstText(result: AgentToolResult<LibreOfficeConvertToolDetails>): string {
  const first = result.content[0];
  if (!first || first.type !== "text") {
    throw new Error("Expected first content block to be text");
  }
  return first.text;
}

void test("libreoffice_convert returns guidance when bridge URL is missing", async () => {
  const tool = createLibreOfficeConvertTool({
    getBridgeConfig: () => Promise.resolve(null),
  });

  const result = await tool.execute("tc-missing", {
    input_path: "/tmp/source.xlsx",
    target_format: "csv",
  });

  assert.match(firstText(result), /native Python bridge/u);
  assert.equal(result.details?.kind, "libreoffice_bridge");
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "missing_bridge_url");
});

void test("libreoffice_convert validates absolute input_path", async () => {
  const tool = createLibreOfficeConvertTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340" }),
    callBridge: () => Promise.resolve({ ok: true, action: "convert" }),
  });

  const result = await tool.execute("tc-validate", {
    input_path: "relative/path.xlsx",
    target_format: "csv",
  });

  assert.match(firstText(result), /input_path must be an absolute path/u);
  assert.equal(result.details?.ok, false);
});

void test("libreoffice_convert sends bridge payload and renders conversion summary", async () => {
  let capturedRequest: LibreOfficeConvertRequest | null = null;
  let capturedConfig: LibreOfficeBridgeConfig | null = null;

  const tool = createLibreOfficeConvertTool({
    getBridgeConfig: () => Promise.resolve({
      url: "https://localhost:3340",
      token: "secret-token",
    }),
    callBridge: (
      request: LibreOfficeConvertRequest,
      config: LibreOfficeBridgeConfig,
    ): Promise<LibreOfficeConvertResponse> => {
      capturedRequest = request;
      capturedConfig = config;

      return Promise.resolve({
        ok: true,
        action: "convert",
        input_path: "/tmp/source.xlsx",
        target_format: "pdf",
        output_path: "/tmp/source.pdf",
        bytes: 4096,
        converter: "soffice",
      });
    },
  });

  const result = await tool.execute("tc-contract", {
    input_path: "/tmp/source.xlsx",
    target_format: "pdf",
    output_path: "/tmp/source.pdf",
    overwrite: true,
    timeout_ms: 12000,
  });

  assert.ok(capturedRequest);
  assert.equal(capturedRequest?.input_path, "/tmp/source.xlsx");
  assert.equal(capturedRequest?.target_format, "pdf");
  assert.equal(capturedRequest?.output_path, "/tmp/source.pdf");
  assert.equal(capturedRequest?.overwrite, true);
  assert.equal(capturedRequest?.timeout_ms, 12000);

  assert.ok(capturedConfig);
  assert.equal(capturedConfig?.url, "https://localhost:3340");
  assert.equal(capturedConfig?.token, "secret-token");

  const text = firstText(result);
  assert.match(text, /Converted/u);
  assert.match(text, /source\.pdf/u);
  assert.match(text, /Output size: 4096 bytes/u);

  assert.equal(result.details?.ok, true);
  assert.equal(result.details?.action, "convert");
  assert.equal(result.details?.targetFormat, "pdf");
  assert.equal(result.details?.outputPath, "/tmp/source.pdf");
});

void test("libreoffice_convert bridge errors are surfaced", async () => {
  const tool = createLibreOfficeConvertTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340" }),
    callBridge: () => Promise.reject(new Error("bridge unavailable")),
  });

  const result = await tool.execute("tc-error", {
    input_path: "/tmp/source.xlsx",
    target_format: "csv",
  });

  assert.equal(firstText(result), "Error: bridge unavailable");
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "bridge unavailable");
});

void test("libreoffice_convert handles explicit bridge rejection payloads", async () => {
  const tool = createLibreOfficeConvertTool({
    getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3340" }),
    callBridge: () => Promise.resolve({
      ok: false,
      action: "convert",
      error: "converter not available",
    }),
  });

  const result = await tool.execute("tc-reject", {
    input_path: "/tmp/source.xlsx",
    target_format: "csv",
  });

  assert.equal(firstText(result), "Error: converter not available");
  assert.equal(result.details?.ok, false);
  assert.equal(result.details?.error, "converter not available");
});
