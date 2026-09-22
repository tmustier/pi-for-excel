import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";
import type {
  Api,
  ImageContent,
  Model,
  ModelImageResizeOptions,
  TextContent,
} from "@earendil-works/pi-ai";

import type { SettingsReader } from "../storage/local/settings-store.js";
import type { Attachment } from "./attachments.js";
import type {
  ImageWorkerRequest,
  ImageWorkerResponse,
} from "./image-input-worker.js";

/** Pi 0.87's conservative cache-safe image resize profile. */
export const DEFAULT_IMAGE_RESIZE_OPTIONS = {
  maxWidth: 2_000,
  maxHeight: 2_000,
  maxBytes: 4.5 * 1024 * 1024,
  jpegQuality: 80,
} as const satisfies Required<ModelImageResizeOptions>;

export const IMAGE_AUTO_RESIZE_SETTING_KEY = "images.autoResize";

export type ProcessedImageInput = {
  ok: true;
  image: ImageContent;
  hints: string[];
} | {
  ok: false;
  message: string;
};

type ImageLimitModel = Pick<Model<Api>, "inputLimits">;

type ImageProcessingOptions = {
  autoResizeImages?: boolean;
  resizeOptions?: ModelImageResizeOptions;
  signal?: AbortSignal;
};

export type ImageInputProcessor = (
  image: ImageContent,
  options?: ImageProcessingOptions,
) => Promise<ProcessedImageInput>;

interface ImageNormalizationOptions {
  autoResizeImages?: boolean;
  processor?: ImageInputProcessor;
  signal?: AbortSignal;
}

interface InstalledImageNormalizationOptions {
  getAutoResizeImages?: () => boolean;
  processor?: ImageInputProcessor;
}

function normalizedMimeType(mimeType: string): ImageContent["mimeType"] | null {
  const baseType = mimeType.split(";")[0]?.trim().toLowerCase();
  if (baseType === "image/png" || baseType === "image/gif" || baseType === "image/webp") {
    return baseType;
  }
  if (baseType === "image/jpeg" || baseType === "image/jpg") {
    return "image/jpeg";
  }
  return null;
}

function baseMimeType(mimeType: string): string {
  return mimeType.split(";")[0]?.trim().toLowerCase() ?? mimeType.toLowerCase();
}

function createAbortError(): Error {
  return new DOMException("Image processing aborted.", "AbortError");
}

function runImageWorker(
  request: ImageWorkerRequest,
  signal?: AbortSignal,
): Promise<ImageWorkerResponse> {
  if (signal?.aborted) return Promise.reject(createAbortError());
  if (typeof Worker === "undefined") {
    return Promise.resolve({ ok: false, reason: "decode" });
  }

  return new Promise<ImageWorkerResponse>((resolve, reject) => {
    let settled = false;
    let worker: Worker;
    try {
      worker = new Worker(
        new URL("./image-input-worker.ts", import.meta.url),
        { type: "module" },
      );
    } catch {
      resolve({ ok: false, reason: "decode" });
      return;
    }

    const cleanup = (): void => {
      signal?.removeEventListener("abort", onAbort);
      worker.terminate();
    };
    const settle = (response: ImageWorkerResponse): void => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(response);
    };
    const fail = (error: Error): void => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(error);
    };
    const onAbort = (): void => fail(createAbortError());

    worker.addEventListener("message", (event: MessageEvent<ImageWorkerResponse>) => {
      settle(event.data);
    }, { once: true });
    worker.addEventListener("error", () => {
      settle({ ok: false, reason: "decode" });
    }, { once: true });
    signal?.addEventListener("abort", onAbort, { once: true });
    worker.postMessage(request);
  });
}

function dimensionHint(
  originalWidth: number,
  originalHeight: number,
  width: number,
  height: number,
): string {
  const scale = originalWidth / width;
  return (
    `[Image: original ${originalWidth}x${originalHeight}, displayed at ${width}x${height}. `
    + `Multiply coordinates by ${scale.toFixed(2)} to map to original image.]`
  );
}

/**
 * Normalize one newly attached image before it enters Agent history.
 *
 * Decode, resize, candidate encoding and base64 conversion run in a dedicated
 * browser worker. Only the accepted candidate crosses back to the Office
 * WebView thread, and aborting terminates the worker immediately.
 */
export async function processImageInput(
  image: ImageContent,
  options: ImageProcessingOptions = {},
): Promise<ProcessedImageInput> {
  if (options.signal?.aborted) throw createAbortError();

  const autoResizeImages = options.autoResizeImages ?? true;
  const supportedMimeType = normalizedMimeType(image.mimeType);

  if (!autoResizeImages && supportedMimeType) {
    if (supportedMimeType === image.mimeType) {
      return { ok: true, image, hints: [] };
    }
    return {
      ok: true,
      image: { ...image, mimeType: supportedMimeType },
      hints: [],
    };
  }

  const limits = { ...DEFAULT_IMAGE_RESIZE_OPTIONS, ...options.resizeOptions };
  const result = await runImageWorker({
    data: image.data,
    mimeType: image.mimeType,
    autoResizeImages,
    canPassThrough: supportedMimeType !== null,
    limits,
  }, options.signal);

  if (!result.ok) {
    return result.reason === "size"
      ? {
          ok: false,
          message: "[Image omitted: could not be resized below the inline image size limit.]",
        }
      : {
          ok: false,
          message: "[Image omitted: could not be converted to a supported inline image format.]",
        };
  }

  if (result.unchanged) {
    if (supportedMimeType === image.mimeType) {
      return { ok: true, image, hints: [] };
    }
    if (supportedMimeType) {
      return {
        ok: true,
        image: { ...image, mimeType: supportedMimeType },
        hints: [],
      };
    }
  }

  if (result.unchanged) {
    return {
      ok: false,
      message: "[Image omitted: could not be converted to a supported inline image format.]",
    };
  }

  const hints: string[] = [];
  if (!supportedMimeType) {
    hints.push(`[Image converted from ${baseMimeType(image.mimeType)} to ${result.image.mimeType}.]`);
  }
  if (result.image.wasResized) {
    hints.push(dimensionHint(
      result.image.originalWidth,
      result.image.originalHeight,
      result.image.width,
      result.image.height,
    ));
  }

  return {
    ok: true,
    image: {
      type: "image",
      data: result.image.data,
      mimeType: result.image.mimeType,
    },
    hints,
  };
}

export async function readImageAutoResizeEnabled(
  settings: SettingsReader,
): Promise<boolean> {
  try {
    const value = await settings.get(IMAGE_AUTO_RESIZE_SETTING_KEY);
    return typeof value === "boolean" ? value : true;
  } catch {
    return true;
  }
}

async function processWithModel(
  image: ImageContent,
  model: ImageLimitModel,
  options: ImageNormalizationOptions,
): Promise<ProcessedImageInput> {
  const processor = options.processor ?? processImageInput;
  const resizeOptions = model.inputLimits?.images?.resize;
  return await processor(image, {
    ...(options.autoResizeImages !== undefined
      ? { autoResizeImages: options.autoResizeImages }
      : {}),
    ...(resizeOptions !== undefined ? { resizeOptions } : {}),
    ...(options.signal !== undefined ? { signal: options.signal } : {}),
  });
}

export async function normalizePromptImages(
  images: readonly ImageContent[] | undefined,
  model: ImageLimitModel,
  options: ImageNormalizationOptions = {},
): Promise<{ images: ImageContent[]; hints: string[] }> {
  if (!images) return { images: [], hints: [] };

  const normalizedImages: ImageContent[] = [];
  const hints: string[] = [];

  for (const image of images) {
    const processed = await processWithModel(image, model, options);
    if (!processed.ok) {
      hints.push(processed.message);
      continue;
    }
    normalizedImages.push(processed.image);
    hints.push(...processed.hints);
  }

  return { images: normalizedImages, hints };
}

async function normalizePromptContent(
  content: string | (TextContent | ImageContent)[],
  model: ImageLimitModel,
  options: ImageNormalizationOptions,
): Promise<string | (TextContent | ImageContent)[]> {
  if (typeof content === "string" || !content.some((block) => block.type === "image")) {
    return content;
  }

  const normalized: (TextContent | ImageContent)[] = [];
  for (const block of content) {
    if (block.type !== "image") {
      normalized.push(block);
      continue;
    }

    const processed = await processWithModel(block, model, options);
    if (!processed.ok) {
      normalized.push({ type: "text", text: processed.message });
      continue;
    }
    if (processed.hints.length > 0) {
      normalized.push({ type: "text", text: processed.hints.join("\n") });
    }
    normalized.push(processed.image);
  }
  return normalized;
}

function appendHints(
  content: string | (TextContent | ImageContent)[],
  hints: readonly string[],
): string | (TextContent | ImageContent)[] {
  if (hints.length === 0) return content;
  const hintText = hints.join("\n");
  if (typeof content === "string") {
    return content.length > 0 ? `${content}\n\n${hintText}` : hintText;
  }
  return [...content, { type: "text", text: hintText }];
}

async function normalizeAttachments(
  attachments: readonly Attachment[] | undefined,
  model: ImageLimitModel,
  options: ImageNormalizationOptions,
): Promise<{ attachments: Attachment[] | undefined; hints: string[] }> {
  if (!attachments?.some((attachment) => attachment.type === "image")) {
    return { attachments: attachments ? [...attachments] : undefined, hints: [] };
  }

  const normalized: Attachment[] = [];
  const hints: string[] = [];
  for (const attachment of attachments) {
    if (attachment.type !== "image") {
      normalized.push(attachment);
      continue;
    }

    const processed = await processWithModel({
      type: "image",
      data: attachment.content,
      mimeType: attachment.mimeType,
    }, model, options);
    if (!processed.ok) {
      hints.push(processed.message);
      continue;
    }

    normalized.push({
      ...attachment,
      content: processed.image.data,
      mimeType: processed.image.mimeType,
    });
    hints.push(...processed.hints);
  }

  return { attachments: normalized, hints };
}

async function normalizePromptMessageInPlace(
  message: AgentMessage,
  model: ImageLimitModel,
  options: ImageNormalizationOptions,
): Promise<void> {
  if (message.role === "user") {
    message.content = await normalizePromptContent(message.content, model, options);
    return;
  }
  if (message.role !== "user-with-attachments" || message.imageInputsNormalized === true) {
    return;
  }

  const normalizedContent = await normalizePromptContent(message.content, model, options);
  const normalizedAttachments = await normalizeAttachments(message.attachments, model, options);
  message.content = appendHints(normalizedContent, normalizedAttachments.hints);
  if (normalizedAttachments.attachments) {
    message.attachments = normalizedAttachments.attachments;
  } else {
    delete message.attachments;
  }
  message.imageInputsNormalized = true;
}

/** Normalize image-bearing prompt messages at the low-level Agent API boundary. */
export function installPromptImageNormalization(
  agent: Agent,
  options: InstalledImageNormalizationOptions = {},
): () => void {
  const normalize = async (message: AgentMessage, signal?: AbortSignal): Promise<void> => {
    await normalizePromptMessageInPlace(message, agent.state.model, {
      autoResizeImages: options.getAutoResizeImages?.() ?? true,
      ...(options.processor !== undefined ? { processor: options.processor } : {}),
      ...(signal !== undefined ? { signal } : {}),
    });
  };

  const unsubscribe = agent.subscribe(async (event, signal) => {
    if (event.type !== "message_end") return;
    if (event.message.role !== "user" && event.message.role !== "user-with-attachments") return;
    await normalize(event.message, signal);
  });

  const previousTransform = agent.transformContext;
  const transform: NonNullable<Agent["transformContext"]> = async (messages, signal) => {
    for (const message of messages) {
      if (message.role === "user-with-attachments" && message.imageInputsNormalized !== true) {
        await normalize(message, signal);
      }
    }
    return previousTransform ? await previousTransform(messages, signal) : messages;
  };
  agent.transformContext = transform;

  return () => {
    unsubscribe();
    if (agent.transformContext !== transform) return;
    if (previousTransform) {
      agent.transformContext = previousTransform;
    } else {
      delete agent.transformContext;
    }
  };
}

export async function normalizeToolResultImages(
  content: (TextContent | ImageContent)[],
  model: ImageLimitModel,
  options: ImageNormalizationOptions = {},
): Promise<(TextContent | ImageContent)[]> {
  if (!content.some((block) => block.type === "image")) return content;

  const normalized: (TextContent | ImageContent)[] = [];
  let changed = false;

  for (const block of content) {
    if (block.type !== "image") {
      normalized.push(block);
      continue;
    }

    const processed = await processWithModel(block, model, options);
    if (!processed.ok) {
      // Pi preserves tool images when the resize backend cannot process them.
      normalized.push(block);
      continue;
    }

    if (
      processed.image.data === block.data
      && processed.image.mimeType === block.mimeType
      && processed.hints.length === 0
    ) {
      normalized.push(block);
      continue;
    }

    normalized.push(processed.image);
    if (processed.hints.length > 0) {
      normalized.push({ type: "text", text: processed.hints.join("\n") });
    }
    changed = true;
  }

  return changed ? normalized : content;
}

/** Install one-time tool-result image normalization on a low-level Agent. */
export function installToolResultImageNormalization(
  agent: Agent,
  options: InstalledImageNormalizationOptions = {},
): () => void {
  const previous = agent.afterToolCall;
  const hook: Agent["afterToolCall"] = async (context, signal) => {
    const previousResult = await previous?.(context, signal);
    const content = previousResult?.content ?? context.result.content;
    const normalized = await normalizeToolResultImages(content, agent.state.model, {
      autoResizeImages: options.getAutoResizeImages?.() ?? true,
      ...(options.processor !== undefined ? { processor: options.processor } : {}),
      ...(signal !== undefined ? { signal } : {}),
    });

    if (normalized.length === content.length && normalized.every((block, index) => block === content[index])) {
      return previousResult;
    }

    return { ...previousResult, content: normalized };
  };

  agent.afterToolCall = hook;
  return () => {
    if (agent.afterToolCall !== hook) return;
    if (previous) {
      agent.afterToolCall = previous;
    } else {
      delete agent.afterToolCall;
    }
  };
}
