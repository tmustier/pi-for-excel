interface ImageWorkerLimits {
  maxWidth: number;
  maxHeight: number;
  maxBytes: number;
  jpegQuality: number;
}

export interface ImageWorkerRequest {
  data: string;
  mimeType: string;
  autoResizeImages: boolean;
  canPassThrough: boolean;
  limits: ImageWorkerLimits;
}

interface WorkerImageResult {
  data: string;
  mimeType: "image/png" | "image/jpeg";
  originalWidth: number;
  originalHeight: number;
  width: number;
  height: number;
  wasResized: boolean;
}

export type ImageWorkerResponse =
  | { ok: true; unchanged: true; originalWidth: number; originalHeight: number }
  | { ok: true; unchanged: false; image: WorkerImageResult }
  | { ok: false; reason: "decode" | "size" };

declare const self: {
  addEventListener(
    type: "message",
    listener: (event: MessageEvent<ImageWorkerRequest>) => void,
  ): void;
  postMessage(message: ImageWorkerResponse): void;
};

function base64ToBytes(data: string): Uint8Array<ArrayBuffer> | null {
  try {
    const binary = atob(data.trim());
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  } catch {
    return null;
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  const chunks: string[] = [];
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + chunkSize)));
  }
  return btoa(chunks.join(""));
}

function encodedBase64Size(blob: Blob): number {
  return Math.ceil(blob.size / 3) * 4;
}

function fitDimensions(
  width: number,
  height: number,
  limits: ImageWorkerLimits,
): { width: number; height: number } {
  let targetWidth = width;
  let targetHeight = height;

  if (targetWidth > limits.maxWidth) {
    targetHeight = Math.max(1, Math.round((targetHeight * limits.maxWidth) / targetWidth));
    targetWidth = limits.maxWidth;
  }
  if (targetHeight > limits.maxHeight) {
    targetWidth = Math.max(1, Math.round((targetWidth * limits.maxHeight) / targetHeight));
    targetHeight = limits.maxHeight;
  }

  return { width: targetWidth, height: targetHeight };
}

async function encodeCandidate(
  canvas: OffscreenCanvas,
  mimeType: "image/png" | "image/jpeg",
  maxBytes: number,
  quality?: number,
): Promise<{ data: string; mimeType: "image/png" | "image/jpeg" } | null> {
  const blob = await canvas.convertToBlob({
    type: mimeType,
    ...(quality !== undefined ? { quality } : {}),
  });
  if (encodedBase64Size(blob) >= maxBytes) return null;

  const bytes = new Uint8Array(await blob.arrayBuffer());
  return { data: bytesToBase64(bytes), mimeType };
}

async function encodeAtDimensions(
  bitmap: ImageBitmap,
  width: number,
  height: number,
  limits: ImageWorkerLimits,
): Promise<{ data: string; mimeType: "image/png" | "image/jpeg" } | null> {
  const canvas = new OffscreenCanvas(width, height);
  const context = canvas.getContext("2d");
  if (!context) return null;
  context.drawImage(bitmap, 0, 0, width, height);

  const png = await encodeCandidate(canvas, "image/png", limits.maxBytes);
  if (png) return png;

  const qualities = Array.from(new Set([limits.jpegQuality, 85, 70, 55, 40]));
  for (const quality of qualities) {
    const jpeg = await encodeCandidate(canvas, "image/jpeg", limits.maxBytes, quality / 100);
    if (jpeg) return jpeg;
  }

  return null;
}

async function processRequest(request: ImageWorkerRequest): Promise<ImageWorkerResponse> {
  if (typeof OffscreenCanvas === "undefined" || typeof createImageBitmap === "undefined") {
    return { ok: false, reason: "decode" };
  }

  const bytes = base64ToBytes(request.data);
  if (!bytes) return { ok: false, reason: "decode" };

  let bitmap: ImageBitmap;
  try {
    bitmap = await createImageBitmap(
      new Blob([bytes], { type: request.mimeType }),
      { imageOrientation: "from-image" },
    );
  } catch {
    return { ok: false, reason: "decode" };
  }

  try {
    const originalWidth = bitmap.width;
    const originalHeight = bitmap.height;
    if (originalWidth < 1 || originalHeight < 1) {
      return { ok: false, reason: "decode" };
    }

    if (!request.autoResizeImages) {
      const converted = await encodeAtDimensions(bitmap, originalWidth, originalHeight, {
        ...request.limits,
        maxBytes: Number.POSITIVE_INFINITY,
      });
      if (!converted) return { ok: false, reason: "decode" };
      return {
        ok: true,
        unchanged: false,
        image: {
          ...converted,
          originalWidth,
          originalHeight,
          width: originalWidth,
          height: originalHeight,
          wasResized: false,
        },
      };
    }

    const inputBase64Size = Math.ceil(bytes.byteLength / 3) * 4;
    if (
      request.canPassThrough
      && originalWidth <= request.limits.maxWidth
      && originalHeight <= request.limits.maxHeight
      && inputBase64Size < request.limits.maxBytes
    ) {
      return { ok: true, unchanged: true, originalWidth, originalHeight };
    }

    let dimensions = fitDimensions(originalWidth, originalHeight, request.limits);
    while (true) {
      const encoded = await encodeAtDimensions(
        bitmap,
        dimensions.width,
        dimensions.height,
        request.limits,
      );
      if (encoded) {
        return {
          ok: true,
          unchanged: false,
          image: {
            ...encoded,
            originalWidth,
            originalHeight,
            width: dimensions.width,
            height: dimensions.height,
            wasResized: true,
          },
        };
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

    return { ok: false, reason: "size" };
  } catch {
    return { ok: false, reason: "decode" };
  } finally {
    bitmap.close();
  }
}

self.addEventListener("message", (event) => {
  void processRequest(event.data).then((response) => {
    self.postMessage(response);
  });
});
