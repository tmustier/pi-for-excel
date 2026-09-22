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

import type { UserMessageWithAttachments } from "../src/messages/attachments.ts";
import { createConvertToLlm } from "../src/messages/convert-to-llm.ts";
import {
  DEFAULT_IMAGE_RESIZE_OPTIONS,
  IMAGE_AUTO_RESIZE_SETTING_KEY,
  installPromptImageNormalization,
  installToolResultImageNormalization,
  normalizePromptImages,
  normalizeToolResultImages,
  readImageAutoResizeEnabled,
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

void test("image auto-resize setting defaults on and accepts only persisted booleans", async () => {
  const read = async (value: unknown, shouldThrow = false): Promise<boolean> => {
    return await readImageAutoResizeEnabled({
      get: (key) => {
        assert.equal(key, IMAGE_AUTO_RESIZE_SETTING_KEY);
        if (shouldThrow) return Promise.reject(new Error("storage unavailable"));
        return Promise.resolve(value);
      },
    });
  };

  assert.equal(await read(undefined), true);
  assert.equal(await read("false"), true);
  assert.equal(await read(false), false);
  assert.equal(await read(true), true);
  assert.equal(await read(undefined, true), true);
});

void test("prompt normalization uses model limits, host resize setting, and AbortSignal", async () => {
  const abortController = new AbortController();
  const calls: Array<{
    autoResizeImages: boolean | undefined;
    maxWidth: number | undefined;
    signal: AbortSignal | undefined;
  }> = [];
  const processor: ImageInputProcessor = (image, options) => {
    calls.push({
      autoResizeImages: options?.autoResizeImages,
      maxWidth: options?.resizeOptions?.maxWidth,
      signal: options?.signal,
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
    signal: abortController.signal,
  });

  assert.deepEqual(calls, [
    { autoResizeImages: undefined, maxWidth: 456, signal: undefined },
    { autoResizeImages: false, maxWidth: 456, signal: abortController.signal },
  ]);
  assert.equal(enabled.images[0]?.data, "processed-1");
  assert.equal(disabled.images[0]?.data, "processed-2");
});

void test("low-level Agent normalizes user-with-attachments at the real prompt API boundary", async () => {
  const faux = fauxProvider({
    provider: "faux-images",
    models: [{
      id: "image-model",
      input: ["text", "image"],
      inputLimits: modelWithResize(456).inputLimits,
      contextWindow: 32_000,
      maxTokens: 4_096,
    }],
  });
  const models = createModels();
  models.setProvider(faux.provider);
  faux.setResponses([fauxAssistantMessage("done")]);

  const requests: TranscriptContext[] = [];
  const agent = new Agent({
    initialState: {
      model: faux.getModel(),
      messages: [],
      tools: [],
    },
    convertToLlm: createConvertToLlm(),
    streamFn: (model, context, options) => {
      requests.push(structuredClone(context));
      return models.streamSimple(model, context, options);
    },
  });

  const signals: AbortSignal[] = [];
  const processor: ImageInputProcessor = (image, options) => {
    if (options?.signal) signals.push(options.signal);
    assert.equal(options?.autoResizeImages, false);
    assert.equal(options?.resizeOptions?.maxWidth, 456);
    return Promise.resolve({
      ok: true,
      image: { ...image, data: "normalized-attachment", mimeType: "image/jpeg" },
      hints: ["[attachment normalized]"],
    });
  };
  const uninstall = installPromptImageNormalization(agent, {
    getAutoResizeImages: () => false,
    processor,
  });

  const prompt: UserMessageWithAttachments = {
    role: "user-with-attachments",
    content: "inspect this",
    timestamp: 1,
    attachments: [{
      id: "image-1",
      type: "image",
      fileName: "chart.png",
      mimeType: "image/png",
      size: 100,
      content: "original-attachment",
    }],
  };
  await agent.prompt(prompt);

  assert.equal(signals.length, 1);
  assert.equal(signals[0]?.aborted, false);
  const stored = agent.state.messages.find((message) => message.role === "user-with-attachments");
  assert.equal(stored?.role, "user-with-attachments");
  if (stored?.role !== "user-with-attachments") assert.fail("expected stored attachment prompt");
  assert.equal(stored.imageInputsNormalized, true);
  assert.equal(stored.attachments?.[0]?.content, "normalized-attachment");
  assert.equal(stored.attachments?.[0]?.mimeType, "image/jpeg");
  assert.equal(stored.content, "inspect this\n\n[attachment normalized]");

  const sent = requests[0]?.messages.find((message) => message.role === "user");
  assert.equal(sent?.role, "user");
  if (sent?.role !== "user" || typeof sent.content === "string") {
    assert.fail("expected converted user attachment content");
  }
  assert.deepEqual(sent.content, [
    { type: "text", text: "inspect this\n\n[attachment normalized]" },
    { type: "image", data: "normalized-attachment", mimeType: "image/jpeg" },
  ]);

  uninstall();
});

void test("restored user-with-attachments history is normalized once before API replay", async () => {
  const faux = fauxProvider({
    provider: "faux-images",
    models: [{
      id: "image-model",
      input: ["text", "image"],
      inputLimits: modelWithResize(222).inputLimits,
      contextWindow: 32_000,
      maxTokens: 4_096,
    }],
  });
  const models = createModels();
  models.setProvider(faux.provider);
  faux.setResponses([fauxAssistantMessage("continued")]);

  const restored: UserMessageWithAttachments = {
    role: "user-with-attachments",
    content: "restored prompt",
    timestamp: 1,
    attachments: [{
      id: "restored-image",
      type: "image",
      fileName: "restored.png",
      mimeType: "image/png",
      size: 100,
      content: "restored-original",
    }],
  };
  const agent = new Agent({
    initialState: {
      model: faux.getModel(),
      messages: [restored],
      tools: [],
    },
    convertToLlm: createConvertToLlm(),
    streamFn: (model, context, options) => models.streamSimple(model, context, options),
  });

  let calls = 0;
  const uninstall = installPromptImageNormalization(agent, {
    processor: (image, options) => {
      calls += 1;
      assert.equal(options?.resizeOptions?.maxWidth, 222);
      return Promise.resolve({
        ok: true,
        image: { ...image, data: "restored-normalized" },
        hints: [],
      });
    },
  });

  await agent.continue();
  assert.equal(calls, 1);
  assert.equal(restored.imageInputsNormalized, true);
  assert.equal(restored.attachments?.[0]?.content, "restored-normalized");

  faux.setResponses([fauxAssistantMessage("continued again")]);
  agent.state.messages = [restored];
  await agent.continue();
  assert.equal(calls, 1);

  uninstall();
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
  const seenSignals: AbortSignal[] = [];
  const processor: ImageInputProcessor = (image, options) => {
    const maxWidth = options?.resizeOptions?.maxWidth;
    if (maxWidth !== undefined) seenMaxWidths.push(maxWidth);
    if (options?.signal) seenSignals.push(options.signal);
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
  assert.equal(seenSignals.length, 1);
  assert.equal(seenSignals[0]?.aborted, false);

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
