import assert from "node:assert/strict";
import { test } from "node:test";
import type { ImageContent } from "@earendil-works/pi-ai";
import type { UserMessageWithAttachments } from "../src/messages/attachments.ts";
import {
  DEFAULT_IMAGE_RESIZE_OPTIONS, normalizePromptContent, normalizePromptMessage,
  normalizeToolResultImages, processImageInput,
} from "../src/messages/image-input-limits.ts";

const image: ImageContent = { type: "image", data: "original", mimeType: "image/png" };
const model = { inputLimits: { images: { resize: { maxWidth: 456 } } } };

void test("Pi conservative image defaults remain exact", () => {
  assert.deepEqual(DEFAULT_IMAGE_RESIZE_OPTIONS, {
    maxWidth: 2000, maxHeight: 2000, maxBytes: 4718592, jpegQuality: 80,
  });
});

void test("prompt content normalization preserves order, forwards limits and omits failed images", async () => {
  const signal = new AbortController().signal;
  let calls = 0;
  const content = await normalizePromptContent([{ type: "text", text: "inspect" }, image, image], model, {
    signal,
    processor: (input, options) => {
      assert.equal(options?.signal, signal);
      assert.equal(options?.resizeOptions?.maxWidth, 456);
      assert.equal(options?.autoResizeImages, true);
      calls++;
      return Promise.resolve(calls === 1
        ? { ok: true, image: { ...input, data: "normalized" }, hints: ["[resized]"] }
        : { ok: false, message: "[omitted]" });
    },
  });
  assert.deepEqual(content, [
    { type: "text", text: "inspect" }, { type: "text", text: "[resized]" },
    { ...image, data: "normalized" }, { type: "text", text: "[omitted]" },
  ]);
  assert.equal(image.data, "original");
  assert.equal(await normalizePromptContent("plain", model, {}), "plain");
});

void test("attachment normalization remains callable and idempotent before submission", async () => {
  const message: UserMessageWithAttachments = {
    role: "user-with-attachments", content: "inspect", timestamp: 1,
    attachments: [{ id: "one", type: "image", fileName: "chart.png", mimeType: "image/png", size: 10, content: "raw" }],
  };
  let calls = 0;
  const options = { processor: (input: ImageContent) => {
    calls++;
    return Promise.resolve({ ok: true as const, image: { ...input, data: "normalized" }, hints: ["[resized]"] });
  } };
  await normalizePromptMessage(message, model, options);
  await normalizePromptMessage(message, model, options);
  assert.equal(calls, 1);
  assert.equal(message.attachments?.[0]?.content, "normalized");
  assert.equal(message.content, "inspect\n\n[resized]");
  assert.equal(message.imageInputsNormalized, true);
});

void test("tool-result normalization preserves failures and adds hints after changed image", async () => {
  const content = [image];
  assert.equal(await normalizeToolResultImages(content, model, {
    processor: () => Promise.resolve({ ok: false, message: "omitted" }),
  }), content);
  assert.deepEqual(await normalizeToolResultImages(content, model, {
    processor: (input) => Promise.resolve({ ok: true, image: { ...input, data: "changed" }, hints: ["hint"] }),
  }), [{ ...image, data: "changed" }, { type: "text", text: "hint" }]);
});

void test("image processing respects cancellation and pass-through MIME canonicalization", async () => {
  const controller = new AbortController();
  controller.abort();
  await assert.rejects(processImageInput(image, { signal: controller.signal }), { name: "AbortError" });
  assert.deepEqual(await processImageInput({ ...image, mimeType: "image/jpg" }, { autoResizeImages: false }), {
    ok: true, image: { ...image, mimeType: "image/jpeg" }, hints: [],
  });
});
