import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";
import type {
  Api,
  ImageContent,
  Model,
  ModelImageResizeOptions,
  TextContent,
} from "@earendil-works/pi-ai";

import type { Attachment } from "./attachments.js";
import type {
  ImageWorkerRequest,
  ImageWorkerResponse,
} from "./image-input-worker.js";

export const DEFAULT_IMAGE_RESIZE_OPTIONS = {
  maxWidth: 2_000,
  maxHeight: 2_000,
  maxBytes: 4.5 * 1024 * 1024,
  jpegQuality: 80,
} as const satisfies Required<ModelImageResizeOptions>;

export type ProcessedImageInput = {
  ok: true;
  image: ImageContent;
  hints: string[];
} | {
  ok: false;
  message: string;
};

type ImageLimitModel = Pick<Model<Api>, "inputLimits">;
type SupportedImageMimeType = NonNullable<ImageWorkerRequest["passThroughMimeType"]>;

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

function createAbortError(): DOMException {
  return new DOMException("Image processing aborted.", "AbortError");
}

function runImageWorker(
  request: ImageWorkerRequest,
  signal?: AbortSignal,
): Promise<ImageWorkerResponse> {
  if (signal?.aborted) return Promise.reject(createAbortError());
  if (typeof Worker === "undefined") return Promise.resolve({ ok: false, reason: "decode" });

  return new Promise<ImageWorkerResponse>((resolve, reject) => {
    let worker: Worker;
    try {
      worker = new Worker(new URL("./image-input-worker.ts", import.meta.url), { type: "module" });
    } catch {
      resolve({ ok: false, reason: "decode" });
      return;
    }

    let settled = false;
    const finish = (response: ImageWorkerResponse): void => {
      if (settled) return;
      settled = true;
      signal?.removeEventListener("abort", abort);
      worker.terminate();
      resolve(response);
    };
    const abort = (): void => {
      if (settled) return;
      settled = true;
      worker.terminate();
      reject(createAbortError());
    };

    worker.onmessage = (event: MessageEvent<ImageWorkerResponse>) => finish(event.data);
    worker.onerror = () => finish({ ok: false, reason: "decode" });
    signal?.addEventListener("abort", abort, { once: true });
    worker.postMessage(request);
  });
}

/** Normalize one image before it is sent to the model or retained in replayable history. */
export async function processImageInput(
  image: ImageContent,
  options: ImageProcessingOptions = {},
): Promise<ProcessedImageInput> {
  if (options.signal?.aborted) throw createAbortError();

  const autoResizeImages = options.autoResizeImages ?? true;
  const originalMimeType = image.mimeType.split(";")[0]?.trim().toLowerCase()
    ?? image.mimeType.toLowerCase();
  let passThroughMimeType: SupportedImageMimeType | null = null;
  if (originalMimeType === "image/jpg") {
    passThroughMimeType = "image/jpeg";
  } else if (
    originalMimeType === "image/png"
    || originalMimeType === "image/jpeg"
    || originalMimeType === "image/gif"
    || originalMimeType === "image/webp"
  ) {
    passThroughMimeType = originalMimeType;
  }
  if (!autoResizeImages && passThroughMimeType) {
    return {
      ok: true,
      image: passThroughMimeType === image.mimeType
        ? image
        : { ...image, mimeType: passThroughMimeType },
      hints: [],
    };
  }

  const result = await runImageWorker({
    data: image.data,
    mimeType: image.mimeType,
    autoResizeImages,
    passThroughMimeType,
    limits: { ...DEFAULT_IMAGE_RESIZE_OPTIONS, ...options.resizeOptions },
  }, options.signal);

  if (!result.ok) {
    return {
      ok: false,
      message: result.reason === "size"
        ? "[Image omitted: could not be resized below the inline image size limit.]"
        : "[Image omitted: could not be converted to a supported inline image format.]",
    };
  }

  if (result.unchanged) {
    return {
      ok: true,
      image: result.mimeType === image.mimeType
        ? image
        : { ...image, mimeType: result.mimeType },
      hints: [],
    };
  }

  const hints: string[] = [];
  if (!passThroughMimeType && originalMimeType) {
    hints.push(`[Image converted from ${originalMimeType} to ${result.image.mimeType}.]`);
  }
  if (result.image.wasResized) {
    const scale = result.image.originalWidth / result.image.width;
    hints.push(
      `[Image: original ${result.image.originalWidth}x${result.image.originalHeight}, `
      + `displayed at ${result.image.width}x${result.image.height}. `
      + `Multiply coordinates by ${scale.toFixed(2)} to map to original image.]`,
    );
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

function processWithModel(
  image: ImageContent,
  model: ImageLimitModel,
  options: ImageNormalizationOptions,
): Promise<ProcessedImageInput> {
  const resizeOptions = model.inputLimits?.images?.resize;
  return (options.processor ?? processImageInput)(image, {
    autoResizeImages: options.autoResizeImages ?? true,
    ...(resizeOptions !== undefined ? { resizeOptions } : {}),
    ...(options.signal !== undefined ? { signal: options.signal } : {}),
  });
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

async function normalizePromptMessage(
  message: AgentMessage,
  model: ImageLimitModel,
  options: ImageNormalizationOptions,
): Promise<void> {
  if (message.role === "user") {
    message.content = await normalizePromptContent(message.content, model, options);
    return;
  }
  if (message.role !== "user-with-attachments" || message.imageInputsNormalized === true) return;

  message.content = await normalizePromptContent(message.content, model, options);
  const hints: string[] = [];
  if (message.attachments?.some((attachment) => attachment.type === "image")) {
    const attachments: Attachment[] = [];
    for (const attachment of message.attachments) {
      if (attachment.type !== "image") {
        attachments.push(attachment);
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
      attachments.push({
        ...attachment,
        content: processed.image.data,
        mimeType: processed.image.mimeType,
      });
      hints.push(...processed.hints);
    }
    message.attachments = attachments;
  }

  if (hints.length > 0) {
    const hintText = hints.join("\n");
    if (typeof message.content === "string") {
      message.content = message.content.length > 0 ? `${message.content}\n\n${hintText}` : hintText;
    } else {
      message.content = [...message.content, { type: "text", text: hintText }];
    }
  }
  message.imageInputsNormalized = true;
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

/** Install prompt and tool-result normalization at the low-level Agent ingress. */
export function installImageInputNormalization(
  agent: Agent,
  options: Pick<ImageNormalizationOptions, "autoResizeImages" | "processor"> = {},
): () => void {
  const normalizedPromptMessages = new WeakSet<AgentMessage>();
  const normalizedToolResultContent = new WeakSet<(TextContent | ImageContent)[]>();
  const processingOptions = (signal?: AbortSignal): ImageNormalizationOptions => ({
    ...options,
    autoResizeImages: options.autoResizeImages ?? true,
    ...(signal ? { signal } : {}),
  });
  const normalize = async (message: AgentMessage, signal?: AbortSignal): Promise<void> => {
    await normalizePromptMessage(message, agent.state.model, processingOptions(signal));
    normalizedPromptMessages.add(message);
  };

  const unsubscribe = agent.subscribe(async (event, signal) => {
    if (
      event.type !== "message_end"
      || (event.message.role !== "user" && event.message.role !== "user-with-attachments")
    ) {
      return;
    }

    try {
      await normalize(event.message, signal);
    } catch (error) {
      if (signal.aborted) {
        await normalizePromptMessage(event.message, agent.state.model, {
          ...processingOptions(),
          processor: () => Promise.resolve({
            ok: false,
            message: "[Image omitted: processing was aborted.]",
          }),
        });
        normalizedPromptMessages.add(event.message);
      }
      throw error;
    }
  });

  const previousTransform = agent.transformContext;
  agent.transformContext = async (messages, signal) => {
    for (const message of messages) {
      if (
        (message.role === "user" || message.role === "user-with-attachments")
        && !normalizedPromptMessages.has(message)
      ) {
        await normalize(message, signal);
        continue;
      }

      if (message.role === "toolResult" && !normalizedToolResultContent.has(message.content)) {
        const normalized = await normalizeToolResultImages(
          message.content,
          agent.state.model,
          processingOptions(signal),
        );
        if (normalized !== message.content) {
          message.content = normalized;
        }
        normalizedToolResultContent.add(message.content);
      }
    }
    return previousTransform ? await previousTransform(messages, signal) : messages;
  };

  const previousAfterToolCall = agent.afterToolCall;
  agent.afterToolCall = async (context, signal) => {
    const previousResult = await previousAfterToolCall?.(context, signal);
    const content = previousResult?.content ?? context.result.content;
    const normalized = await normalizeToolResultImages(
      content,
      agent.state.model,
      processingOptions(signal),
    );
    normalizedToolResultContent.add(normalized);
    return normalized === content ? previousResult : { ...previousResult, content: normalized };
  };

  return () => {
    unsubscribe();
    if (previousTransform) {
      agent.transformContext = previousTransform;
    } else {
      delete agent.transformContext;
    }
    if (previousAfterToolCall) {
      agent.afterToolCall = previousAfterToolCall;
    } else {
      delete agent.afterToolCall;
    }
  };
}
