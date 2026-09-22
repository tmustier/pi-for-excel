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
  installImageInputNormalization,
  normalizeToolResultImages,
  processImageInput,
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

void test("conversion hints require a non-empty source MIME type", async () => {
  const workerDescriptor = Object.getOwnPropertyDescriptor(globalThis, "Worker");
  Object.defineProperty(globalThis, "Worker", {
    configurable: true,
    value: class {
      onmessage: ((event: MessageEvent) => void) | null = null;

      postMessage(): void {
        this.onmessage?.(new MessageEvent("message", {
          data: {
            ok: true,
            unchanged: false,
            image: {
              data: "converted-base64",
              mimeType: "image/png",
              originalWidth: 1,
              originalHeight: 1,
              width: 1,
              height: 1,
              wasResized: false,
            },
          },
        }));
      }

      terminate(): void {}
    },
  });

  try {
    for (const mimeType of ["", " \t"]) {
      const processed = await processImageInput({ ...IMAGE, mimeType });
      assert.equal(processed.ok, true);
      if (!processed.ok) assert.fail("expected converted image");
      assert.deepEqual(processed.hints, []);
    }
  } finally {
    if (workerDescriptor) {
      Object.defineProperty(globalThis, "Worker", workerDescriptor);
    } else {
      Reflect.deleteProperty(globalThis, "Worker");
    }
  }
});

void test("low-level Agent normalizes image-array prompts at the real prompt API boundary", async () => {
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
    streamFn: (model, context, options) => {
      requests.push(structuredClone(context));
      return models.streamSimple(model, context, options);
    },
  });

  let calls = 0;
  const uninstall = installImageInputNormalization(agent, {
    processor: (image, options) => {
      calls += 1;
      assert.equal(options?.autoResizeImages, true);
      assert.equal(options?.resizeOptions?.maxWidth, 456);
      assert.equal(options?.signal?.aborted, false);
      return Promise.resolve({
        ok: true,
        image: { ...image, data: "normalized-api-image", mimeType: "image/jpeg" },
        hints: ["[api image normalized]"],
      });
    },
  });

  await agent.prompt("inspect this", [IMAGE]);

  assert.equal(calls, 1);
  const stored = agent.state.messages.find((message) => message.role === "user");
  assert.equal(stored?.role, "user");
  if (stored?.role !== "user" || typeof stored.content === "string") {
    assert.fail("expected stored image-array prompt");
  }
  assert.deepEqual(stored.content, [
    { type: "text", text: "inspect this" },
    { type: "text", text: "[api image normalized]" },
    { type: "image", data: "normalized-api-image", mimeType: "image/jpeg" },
  ]);
  assert.deepEqual(requests[0]?.messages[0], stored);

  uninstall();
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
  const uninstall = installImageInputNormalization(agent, {
    autoResizeImages: false,
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
  const uninstall = installImageInputNormalization(agent, {
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
  const uninstall = installImageInputNormalization(agent, { processor });

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

void test("restored tool-result images are normalized once before resumed requests", async () => {
  const faux = fauxProvider({
    provider: "faux-images",
    models: [{
      id: "image-model",
      input: ["text", "image"],
      inputLimits: modelWithResize(900).inputLimits,
      contextWindow: 32_000,
      maxTokens: 4_096,
    }],
  });
  const models = createModels();
  models.setProvider(faux.provider);
  faux.setResponses([
    fauxAssistantMessage("resumed"),
    fauxAssistantMessage("continued"),
  ]);

  const restored = {
    role: "toolResult" as const,
    toolCallId: "restored-call",
    toolName: "charts",
    content: [IMAGE],
    isError: false,
    timestamp: 1,
  };
  const requests: TranscriptContext[] = [];
  const agent = new Agent({
    initialState: {
      model: faux.getModel(),
      messages: [],
      tools: [],
    },
    streamFn: (model, context, options) => {
      requests.push(structuredClone(context));
      return models.streamSimple(model, context, options);
    },
  });

  let calls = 0;
  const uninstall = installImageInputNormalization(agent, {
    processor: (image, options) => {
      calls += 1;
      assert.equal(options?.resizeOptions?.maxWidth, 900);
      return Promise.resolve({
        ok: true,
        image: { ...image, data: "restored-tool-image-normalized" },
        hints: ["[restored tool image normalized]"],
      });
    },
  });

  agent.state.messages = [restored];
  await agent.continue();
  assert.equal(calls, 1);
  assert.equal(restored.content[0]?.type, "image");
  assert.equal(
    restored.content[0]?.type === "image" ? restored.content[0].data : "",
    "restored-tool-image-normalized",
  );
  assert.equal(requests[0]?.messages[0]?.role, "toolResult");
  assert.match(JSON.stringify(requests[0]), /restored-tool-image-normalized/u);
  assert.doesNotMatch(JSON.stringify(requests[0]), /original-base64/u);

  agent.state.model = {
    ...agent.state.model,
    inputLimits: { images: { resize: { maxWidth: 300 } } },
  };
  await agent.prompt("continue");
  assert.equal(calls, 1);
  assert.match(JSON.stringify(requests[1]), /restored-tool-image-normalized/u);

  uninstall();
});

void test("aborting prompt image normalization removes the raw image from replayable history", async () => {
  const faux = fauxProvider({
    provider: "faux-images",
    models: [{
      id: "image-model",
      input: ["text", "image"],
      inputLimits: modelWithResize(700).inputLimits,
      contextWindow: 32_000,
      maxTokens: 4_096,
    }],
  });
  const models = createModels();
  models.setProvider(faux.provider);
  const requests: TranscriptContext[] = [];
  const agent = new Agent({
    initialState: {
      model: faux.getModel(),
      messages: [],
      tools: [],
    },
    streamFn: (model, context, options) => {
      requests.push(structuredClone(context));
      return models.streamSimple(model, context, options);
    },
  });

  let processorStarted = (): void => {};
  const started = new Promise<void>((resolve) => {
    processorStarted = resolve;
  });
  const uninstall = installImageInputNormalization(agent, {
    processor: (_image, options) => new Promise((_resolve, reject) => {
      processorStarted();
      options?.signal?.addEventListener("abort", () => {
        reject(new DOMException("Image processing aborted.", "AbortError"));
      }, { once: true });
    }),
  });

  const prompt = agent.prompt("inspect this", [IMAGE]);
  await started;
  agent.abort();
  await prompt;

  assert.equal(requests.length, 0);
  assert.doesNotMatch(JSON.stringify(agent.state.messages), /original-base64/u);
  const stored = agent.state.messages.find((message) => message.role === "user");
  assert.equal(stored?.role, "user");
  if (stored?.role !== "user" || typeof stored.content === "string") {
    assert.fail("expected sanitized user prompt");
  }
  assert.deepEqual(stored.content, [
    { type: "text", text: "inspect this" },
    { type: "text", text: "[Image omitted: processing was aborted.]" },
  ]);
  const aborted = agent.state.messages.at(-1);
  assert.equal(aborted?.role, "assistant");
  if (aborted?.role !== "assistant") assert.fail("expected aborted assistant message");
  assert.equal(aborted.stopReason, "aborted");

  faux.setResponses([fauxAssistantMessage("continued")]);
  await agent.prompt("continue after abort");
  assert.equal(requests.length, 1);
  assert.doesNotMatch(JSON.stringify(requests[0]), /original-base64/u);

  uninstall();
});
