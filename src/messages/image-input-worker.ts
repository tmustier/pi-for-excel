interface ImageWorkerLimits {
  maxWidth: number;
  maxHeight: number;
  maxBytes: number;
  jpegQuality: number;
}

type PassThroughMimeType = "image/png" | "image/jpeg" | "image/gif" | "image/webp";
type EncodedMimeType = "image/png" | "image/jpeg";

export interface ImageWorkerRequest {
  data: string;
  mimeType: string;
  autoResizeImages: boolean;
  passThroughMimeType: PassThroughMimeType | null;
  limits: ImageWorkerLimits;
}

interface WorkerImageResult {
  data: string;
  mimeType: EncodedMimeType;
  originalWidth: number;
  originalHeight: number;
  width: number;
  height: number;
  wasResized: boolean;
}

export type ImageWorkerResponse =
  | { ok: true; unchanged: true; mimeType: PassThroughMimeType }
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
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    chunks.push(String.fromCharCode(...bytes.subarray(offset, offset + 0x8000)));
  }
  return btoa(chunks.join(""));
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
  mimeType: EncodedMimeType,
  maxBytes: number,
  quality?: number,
): Promise<{ data: string; mimeType: EncodedMimeType } | null> {
  const blob = await canvas.convertToBlob({
    type: mimeType,
    ...(quality !== undefined ? { quality } : {}),
  });
  if (Math.ceil(blob.size / 3) * 4 >= maxBytes) return null;
  return {
    data: bytesToBase64(new Uint8Array(await blob.arrayBuffer())),
    mimeType,
  };
}

async function encodeAtDimensions(
  bitmap: ImageBitmap,
  width: number,
  height: number,
  limits: ImageWorkerLimits,
): Promise<{ data: string; mimeType: EncodedMimeType } | null> {
  const canvas = new OffscreenCanvas(width, height);
  const context = canvas.getContext("2d");
  if (!context) return null;
  context.drawImage(bitmap, 0, 0, width, height);

  const png = await encodeCandidate(canvas, "image/png", limits.maxBytes);
  if (png) return png;
  for (const quality of new Set([limits.jpegQuality, 85, 70, 55, 40])) {
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

    if (!request.autoResizeImages) {
      const converted = await encodeAtDimensions(bitmap, originalWidth, originalHeight, {
        ...request.limits,
        maxBytes: Number.POSITIVE_INFINITY,
      });
      return converted
        ? {
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
          }
        : { ok: false, reason: "decode" };
    }

    const withinLimits = originalWidth <= request.limits.maxWidth
      && originalHeight <= request.limits.maxHeight
      && Math.ceil(bytes.byteLength / 3) * 4 < request.limits.maxBytes;
    if (request.passThroughMimeType && withinLimits) {
      return { ok: true, unchanged: true, mimeType: request.passThroughMimeType };
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
        const wasResized = request.passThroughMimeType !== null
          || dimensions.width !== originalWidth
          || dimensions.height !== originalHeight
          || encoded.mimeType !== "image/png";
        return {
          ok: true,
          unchanged: false,
          image: {
            ...encoded,
            originalWidth,
            originalHeight,
            width: dimensions.width,
            height: dimensions.height,
            wasResized,
          },
        };
      }

      if (dimensions.width === 1 && dimensions.height === 1) break;
      const next = {
        width: dimensions.width === 1 ? 1 : Math.max(1, Math.floor(dimensions.width * 0.75)),
        height: dimensions.height === 1 ? 1 : Math.max(1, Math.floor(dimensions.height * 0.75)),
      };
      if (next.width === dimensions.width && next.height === dimensions.height) break;
      dimensions = next;
    }
    return { ok: false, reason: "size" };
  } catch {
    return { ok: false, reason: "decode" };
  } finally {
    bitmap.close();
  }
}

self.addEventListener("message", (event) => {
  void processRequest(event.data).then((response) => self.postMessage(response));
});
