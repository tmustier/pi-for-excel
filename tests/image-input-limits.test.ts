import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentTool } from "@earendil-works/pi-agent-core";
import {
  createModels,
  fauxAssistantMessage,
  fauxProvider,
  fauxToolCall,
  type Api,
  type ImageContent,
  type Model,
  type TranscriptContext,
} from "@earendil-works/pi-ai";
import { Type } from "typebox";

import {
  DEFAULT_IMAGE_RESIZE_OPTIONS,
  installToolResultImageNormalization,
  normalizePromptImages,
  normalizeToolResultImages,
  type ImageInputProcessor,
} from "../src/messages/image-input-limits.ts";

const IMAGE: ImageContent = {
  type: "image",
  data: "original-base64",
  mimeType: "image/png",
};

function modelWithResize(maxWidth: number): Model<Api> {
  return {
    id: "image-model",
    name: "Image model",
    api: "faux",
    provider: "faux-images",
    baseUrl: "https://example.invalid",
    reasoning: false,
    input: ["text", "image"],
    inputLimits: {
      images: {
        resize: {
          maxWidth,
          maxHeight: 777,
          maxBytes: 1_234,
          jpegQuality: 67,
        },
      },
    },
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow: 32_000,
    maxTokens: 4_096,
  };
}

void test("Pi conservative image defaults remain exact", () => {
  assert.deepEqual(DEFAULT_IMAGE_RESIZE_OPTIONS, {
    maxWidth: 2_000,
    maxHeight: 2_000,
    maxBytes: 4_718_592,
    jpegQuality: 80,
  });
});

void test("prompt normalization uses model limits and honors host resize disable", async () => {
  const calls: Array<{ autoResizeImages: boolean | undefined; maxWidth: number | undefined }> = [];
  const processor: ImageInputProcessor = (image, options) => {
    calls.push({
      autoResizeImages: options?.autoResizeImages,
      maxWidth: options?.resizeOptions?.maxWidth,
    });
    return Promise.resolve({
      ok: true,
      image: { ...image, data: `processed-${calls.length}` },
      hints: [],
    });
  };
  const model = modelWithResize(456);

  const enabled = await normalizePromptImages([IMAGE], model, { processor });
  const disabled = await normalizePromptImages([IMAGE], model, {
    autoResizeImages: false,
    processor,
  });

  assert.deepEqual(calls, [
    { autoResizeImages: undefined, maxWidth: 456 },
    { autoResizeImages: false, maxWidth: 456 },
  ]);
  assert.equal(enabled.images[0]?.data, "processed-1");
  assert.equal(disabled.images[0]?.data, "processed-2");
});

void test("tool-result normalization keeps failed images and appends resize hints only on change", async () => {
  const content = [
    { type: "text" as const, text: "chart" },
    IMAGE,
  ];
  const failed = await normalizeToolResultImages(content, modelWithResize(500), {
    processor: () => Promise.resolve({ ok: false, message: "omitted" }),
  });
  assert.equal(failed, content);

  const changed = await normalizeToolResultImages(content, modelWithResize(500), {
    processor: (image) => Promise.resolve({
      ok: true,
      image: { ...image, data: "resized-base64", mimeType: "image/jpeg" },
      hints: ["[resized]"],
    }),
  });
  assert.deepEqual(changed, [
    content[0],
    { type: "image", data: "resized-base64", mimeType: "image/jpeg" },
    { type: "text", text: "[resized]" },
  ]);
});

void test("low-level Agent normalizes each new tool image once and does not rewrite it after model changes", async () => {
  const faux = fauxProvider({
    provider: "faux-images",
    models: [{
      id: "image-model",
      input: ["text", "image"],
      inputLimits: modelWithResize(640).inputLimits,
      contextWindow: 32_000,
      maxTokens: 4_096,
    }],
  });
  const models = createModels();
  models.setProvider(faux.provider);
  faux.setResponses([
    fauxAssistantMessage(fauxToolCall("capture", {}, { id: "capture-1" }), {
      stopReason: "toolUse",
    }),
    fauxAssistantMessage("first done"),
    fauxAssistantMessage("second done"),
  ]);

  const captureSchema = Type.Object({});
  const captureTool: AgentTool<typeof captureSchema, undefined> = {
    name: "capture",
    label: "capture",
    description: "capture an image",
    parameters: captureSchema,
    execute: () => Promise.resolve({ content: [IMAGE], details: undefined }),
  };
  const requests: TranscriptContext[] = [];
  const agent = new Agent({
    initialState: {
      model: faux.getModel(),
      messages: [],
      tools: [captureTool],
    },
    streamFn: (model, context, options) => {
      requests.push(structuredClone(context));
      return models.streamSimple(model, context, options);
    },
  });

  const seenMaxWidths: number[] = [];
  const processor: ImageInputProcessor = (image, options) => {
    const maxWidth = options?.resizeOptions?.maxWidth;
    if (maxWidth !== undefined) seenMaxWidths.push(maxWidth);
    return Promise.resolve({
      ok: true,
      image: { ...image, data: `resized-for-${String(maxWidth)}` },
      hints: [],
    });
  };
  const uninstall = installToolResultImageNormalization(agent, { processor });

  await agent.prompt("capture it");
  const storedToolResult = agent.state.messages.find((message) => message.role === "toolResult");
  assert.equal(storedToolResult?.role, "toolResult");
  if (storedToolResult?.role !== "toolResult") assert.fail("expected a stored tool result");
  assert.equal(storedToolResult.content[0]?.type, "image");
  assert.equal(storedToolResult.content[0]?.type === "image" ? storedToolResult.content[0].data : "", "resized-for-640");
  assert.deepEqual(seenMaxWidths, [640]);

  agent.state.model = {
    ...agent.state.model,
    inputLimits: { images: { resize: { maxWidth: 320 } } },
  };
  await agent.prompt("continue without another capture");

  assert.deepEqual(seenMaxWidths, [640]);
  const replayed = requests[requests.length - 1]?.messages.find(
    (message) => message.role === "toolResult",
  );
  assert.equal(replayed?.role, "toolResult");
  if (replayed?.role !== "toolResult") assert.fail("expected replayed tool result");
  assert.equal(replayed.content[0]?.type === "image" ? replayed.content[0].data : "", "resized-for-640");

  uninstall();
});
