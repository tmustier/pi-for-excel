import type { Agent } from "@earendil-works/pi-agent-core";
import type {
  Api,
  ImageContent,
  Model,
  ModelImageResizeOptions,
  TextContent,
} from "@earendil-works/pi-ai";

import { bytesToBase64 } from "../files/encoding.js";

/** Pi 0.87's conservative cache-safe image resize profile. */
export const DEFAULT_IMAGE_RESIZE_OPTIONS = {
  maxWidth: 2_000,
  maxHeight: 2_000,
  maxBytes: 4.5 * 1024 * 1024,
  jpegQuality: 80,
} as const satisfies Required<ModelImageResizeOptions>;

interface DecodedImage {
  element: HTMLImageElement;
  width: number;
  height: number;
  release: () => void;
}

export type ProcessedImageInput = {
  ok: true;
  image: ImageContent;
  hints: string[];
} | {
  ok: false;
  message: string;
};

type ImageLimitModel = Pick<Model<Api>, "inputLimits">;

export type ImageInputProcessor = (
  image: ImageContent,
  options?: {
    autoResizeImages?: boolean;
    resizeOptions?: ModelImageResizeOptions;
  },
) => Promise<ProcessedImageInput>;

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

async function decodeImage(image: ImageContent): Promise<DecodedImage | null> {
  let bytes: Uint8Array;
  try {
    const binary = atob(image.data.trim());
    bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
  } catch {
    return null;
  }

  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  const objectUrl = URL.createObjectURL(new Blob([buffer], { type: image.mimeType }));
  const element = new Image();
  try {
    await new Promise<void>((resolve, reject) => {
      element.onload = () => resolve();
      element.onerror = () => reject(new Error("Image decode failed."));
      element.src = objectUrl;
    });
  } catch {
    URL.revokeObjectURL(objectUrl);
    return null;
  }

  if (element.naturalWidth < 1 || element.naturalHeight < 1) {
    URL.revokeObjectURL(objectUrl);
    return null;
  }

  return {
    element,
    width: element.naturalWidth,
    height: element.naturalHeight,
    release: () => URL.revokeObjectURL(objectUrl),
  };
}

function canvasToBlob(
  canvas: HTMLCanvasElement,
  mimeType: "image/png" | "image/jpeg",
  quality?: number,
): Promise<Blob | null> {
  return new Promise((resolve) => canvas.toBlob(resolve, mimeType, quality));
}

async function encodeCandidate(
  canvas: HTMLCanvasElement,
  mimeType: "image/png" | "image/jpeg",
  quality?: number,
): Promise<ImageContent | null> {
  const blob = await canvasToBlob(canvas, mimeType, quality);
  if (!blob) return null;

  return {
    type: "image",
    data: bytesToBase64(new Uint8Array(await blob.arrayBuffer())),
    mimeType,
  };
}

function fitDimensions(
  width: number,
  height: number,
  limits: Required<ModelImageResizeOptions>,
): { width: number; height: number } {
  let targetWidth = width;
  let targetHeight = height;

  if (targetWidth > limits.maxWidth) {
    targetHeight = Math.round((targetHeight * limits.maxWidth) / targetWidth);
    targetWidth = limits.maxWidth;
  }
  if (targetHeight > limits.maxHeight) {
    targetWidth = Math.round((targetWidth * limits.maxHeight) / targetHeight);
    targetHeight = limits.maxHeight;
  }

  return { width: targetWidth, height: targetHeight };
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

function drawToCanvas(
  decoded: DecodedImage,
  width: number,
  height: number,
): HTMLCanvasElement | null {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) return null;
  context.drawImage(decoded.element, 0, 0, width, height);
  return canvas;
}

async function encodeWithinLimits(
  decoded: DecodedImage,
  width: number,
  height: number,
  limits: Required<ModelImageResizeOptions>,
): Promise<ImageContent | null> {
  const canvas = drawToCanvas(decoded, width, height);
  if (!canvas) return null;

  const png = await encodeCandidate(canvas, "image/png");
  if (png && png.data.length < limits.maxBytes) return png;

  const qualities = Array.from(new Set([limits.jpegQuality, 85, 70, 55, 40]));
  for (const quality of qualities) {
    const jpeg = await encodeCandidate(canvas, "image/jpeg", quality / 100);
    if (jpeg && jpeg.data.length < limits.maxBytes) return jpeg;
  }

  return null;
}

/**
 * Normalize one newly attached image before it enters Agent history.
 *
 * The browser implementation mirrors Pi 0.87's dimensions, encoded-size limit,
 * quality sequence and progressive 75% fallback. It deliberately runs only at
 * ingress so later requests keep a stable encoded prefix.
 */
export async function processImageInput(
  image: ImageContent,
  options: {
    autoResizeImages?: boolean;
    resizeOptions?: ModelImageResizeOptions;
  } = {},
): Promise<ProcessedImageInput> {
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

  const decoded = await decodeImage(image);
  if (!decoded) {
    return {
      ok: false,
      message: "[Image omitted: could not be converted to a supported inline image format.]",
    };
  }

  try {
    if (!autoResizeImages) {
      const canvas = drawToCanvas(decoded, decoded.width, decoded.height);
      const converted = canvas ? await encodeCandidate(canvas, "image/png") : null;
      if (!converted) {
        return {
          ok: false,
          message: "[Image omitted: could not be converted to a supported inline image format.]",
        };
      }
      return {
        ok: true,
        image: converted,
        hints: [`[Image converted from ${image.mimeType} to image/png.]`],
      };
    }

    const limits = { ...DEFAULT_IMAGE_RESIZE_OPTIONS, ...options.resizeOptions };
    if (
      supportedMimeType
      && decoded.width <= limits.maxWidth
      && decoded.height <= limits.maxHeight
      && image.data.length < limits.maxBytes
    ) {
      if (supportedMimeType === image.mimeType) {
        return { ok: true, image, hints: [] };
      }
      return {
        ok: true,
        image: { ...image, mimeType: supportedMimeType },
        hints: [],
      };
    }

    let dimensions = fitDimensions(decoded.width, decoded.height, limits);
    while (true) {
      const encoded = await encodeWithinLimits(
        decoded,
        dimensions.width,
        dimensions.height,
        limits,
      );
      if (encoded) {
        const hints: string[] = [];
        if (!supportedMimeType) {
          hints.push(`[Image converted from ${image.mimeType} to ${encoded.mimeType}.]`);
        }
        if (dimensions.width !== decoded.width || dimensions.height !== decoded.height) {
          hints.push(dimensionHint(
            decoded.width,
            decoded.height,
            dimensions.width,
            dimensions.height,
          ));
        }
        return { ok: true, image: encoded, hints };
      }

      if (dimensions.width === 1 && dimensions.height === 1) break;
      const nextWidth = dimensions.width === 1
        ? 1
        : Math.max(1, Math.floor(dimensions.width * 0.75));
      const nextHeight = dimensions.height === 1
        ? 1
        : Math.max(1, Math.floor(dimensions.height * 0.75));
      if (nextWidth === dimensions.width && nextHeight === dimensions.height) break;
      dimensions = { width: nextWidth, height: nextHeight };
    }

    return {
      ok: false,
      message: "[Image omitted: could not be resized below the inline image size limit.]",
    };
  } finally {
    decoded.release();
  }
}

export async function normalizePromptImages(
  images: readonly ImageContent[] | undefined,
  model: ImageLimitModel,
  options: {
    autoResizeImages?: boolean;
    processor?: ImageInputProcessor;
  } = {},
): Promise<{ images: ImageContent[]; hints: string[] }> {
  if (!images) return { images: [], hints: [] };

  const normalizedImages: ImageContent[] = [];
  const hints: string[] = [];
  const processor = options.processor ?? processImageInput;

  for (const image of images) {
    const resizeOptions = model.inputLimits?.images?.resize;
    const processed = await processor(image, {
      ...(options.autoResizeImages !== undefined
        ? { autoResizeImages: options.autoResizeImages }
        : {}),
      ...(resizeOptions !== undefined ? { resizeOptions } : {}),
    });
    if (!processed.ok) {
      hints.push(processed.message);
      continue;
    }
    normalizedImages.push(processed.image);
    hints.push(...processed.hints);
  }

  return { images: normalizedImages, hints };
}

export async function normalizeToolResultImages(
  content: (TextContent | ImageContent)[],
  model: ImageLimitModel,
  options: {
    autoResizeImages?: boolean;
    processor?: ImageInputProcessor;
  } = {},
): Promise<(TextContent | ImageContent)[]> {
  if (!content.some((block) => block.type === "image")) return content;

  const normalized: (TextContent | ImageContent)[] = [];
  const processor = options.processor ?? processImageInput;
  let changed = false;

  for (const block of content) {
    if (block.type !== "image") {
      normalized.push(block);
      continue;
    }

    const resizeOptions = model.inputLimits?.images?.resize;
    const processed = await processor(block, {
      ...(options.autoResizeImages !== undefined
        ? { autoResizeImages: options.autoResizeImages }
        : {}),
      ...(resizeOptions !== undefined ? { resizeOptions } : {}),
    });
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
  options: {
    autoResizeImages?: boolean;
    processor?: ImageInputProcessor;
  } = {},
): () => void {
  const previous = agent.afterToolCall;
  const hook: Agent["afterToolCall"] = async (context, signal) => {
    const previousResult = await previous?.(context, signal);
    const content = previousResult?.content ?? context.result.content;
    const normalized = await normalizeToolResultImages(content, agent.state.model, options);

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
