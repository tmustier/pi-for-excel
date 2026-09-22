import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
let previousAutoResizeConfig: string | undefined;

before(async () => {
  previousAutoResizeConfig = process.env.VITE_PI_AUTO_RESIZE_IMAGES;
  process.env.VITE_PI_AUTO_RESIZE_IMAGES = "false";
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
  if (previousAutoResizeConfig === undefined) {
    delete process.env.VITE_PI_AUTO_RESIZE_IMAGES;
  } else {
    process.env.VITE_PI_AUTO_RESIZE_IMAGES = previousAutoResizeConfig;
  }
});

void test("bootstrap image config reaches the installed tool-result hook", async () => {
  const opened = await openTaskpane(env, { clientId: "image-config-browser-client" });
  try {
    const workerStarts = await opened.page.evaluate(`
      (async () => {
        const sidebar = document.querySelector("pi-sidebar");
        const agent = sidebar?.agent;
        if (!agent?.afterToolCall) throw new Error("image normalization hook was not installed");

        const OriginalWorker = globalThis.Worker;
        let starts = 0;
        globalThis.Worker = class {
          constructor() {
            starts += 1;
            throw new Error("worker should not start when bootstrap disables auto-resize");
          }
        };
        try {
          await agent.afterToolCall({
            result: {
              content: [{ type: "image", data: "invalid-image", mimeType: "image/png" }],
            },
          }, new AbortController().signal);
          return starts;
        } finally {
          globalThis.Worker = OriginalWorker;
        }
      })()
    `);

    assert.equal(workerStarts, 0);
  } finally {
    await opened.finish();
  }
});

void test("browser image worker matches Pi limits without blocking the WebView", async () => {
  const opened = await openTaskpane(env, { clientId: "image-limits-browser-client" });
  try {
    const result = await opened.page.evaluate(`
      (async () => {
        const {
          normalizeToolResultImages,
          processImageInput,
        } = await import("/src/messages/image-input-limits.ts");

        const imageDimensions = (image) => new Promise((resolve, reject) => {
          const decoded = new Image();
          decoded.onload = () => resolve({
            width: decoded.naturalWidth,
            height: decoded.naturalHeight,
          });
          decoded.onerror = () => reject(new Error("processed image did not decode"));
          decoded.src = "data:" + image.mimeType + ";base64," + image.data;
        });
        const canvasBase64 = (width, height, paint) => {
          const canvas = document.createElement("canvas");
          canvas.width = width;
          canvas.height = height;
          const context = canvas.getContext("2d");
          paint(context, canvas);
          return canvas.toDataURL("image/png").split(",")[1];
        };
        const blobBase64Size = async (canvas, quality) => {
          const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", quality));
          return Math.ceil(blob.size / 3) * 4;
        };

        const original = canvasBase64(120, 80, (context, canvas) => {
          context.fillStyle = "#c2185b";
          context.fillRect(0, 0, canvas.width, canvas.height);
        });
        const model = {
          inputLimits: {
            images: {
              resize: {
                maxWidth: 40,
                maxHeight: 40,
                maxBytes: 100000,
                jpegQuality: 80,
              },
            },
          },
        };
        const input = { type: "image", data: original, mimeType: "image/png" };
        const normalized = await processImageInput(input, {
          resizeOptions: model.inputLimits.images.resize,
        });
        const disabled = await processImageInput(input, { autoResizeImages: false });
        const resized = normalized.image;

        const toolResult = await normalizeToolResultImages([
          { type: "text", text: "before" },
          input,
        ], model);

        const noiseCanvas = document.createElement("canvas");
        noiseCanvas.width = 180;
        noiseCanvas.height = 120;
        const noiseContext = noiseCanvas.getContext("2d");
        const pixels = noiseContext.createImageData(noiseCanvas.width, noiseCanvas.height);
        let seed = 123456789;
        for (let index = 0; index < pixels.data.length; index += 4) {
          seed = (seed * 1664525 + 1013904223) >>> 0;
          pixels.data[index] = seed & 255;
          pixels.data[index + 1] = (seed >>> 8) & 255;
          pixels.data[index + 2] = (seed >>> 16) & 255;
          pixels.data[index + 3] = 255;
        }
        noiseContext.putImageData(pixels, 0, 0);
        const noisyPng = noiseCanvas.toDataURL("image/png").split(",")[1];
        const jpeg85Size = await blobBase64Size(noiseCanvas, 0.85);
        const jpeg70Size = await blobBase64Size(noiseCanvas, 0.70);
        const qualityCap = Math.floor((jpeg85Size + jpeg70Size) / 2);
        const byteOnly = await processImageInput({
          type: "image",
          data: noisyPng,
          mimeType: "image/png",
        }, {
          resizeOptions: {
            maxWidth: 1000,
            maxHeight: 1000,
            maxBytes: qualityCap,
            jpegQuality: 90,
          },
        });

        const invalid = await processImageInput({
          type: "image",
          data: "bm90LWFuLWltYWdl",
          mimeType: "image/png",
        });

        const bmp = new Uint8Array(58);
        const bmpView = new DataView(bmp.buffer);
        bmp[0] = 0x42;
        bmp[1] = 0x4d;
        bmpView.setUint32(2, bmp.length, true);
        bmpView.setUint32(10, 54, true);
        bmpView.setUint32(14, 40, true);
        bmpView.setInt32(18, 1, true);
        bmpView.setInt32(22, 1, true);
        bmpView.setUint16(26, 1, true);
        bmpView.setUint16(28, 24, true);
        bmpView.setUint32(34, 4, true);
        bmp[56] = 0xff;
        const bmpInput = {
          type: "image",
          data: btoa(String.fromCharCode(...bmp)),
          mimeType: "image/bmp",
        };
        const converted = await processImageInput(bmpInput, { autoResizeImages: false });
        const convertedWithResize = await processImageInput(bmpInput);

        const orientedJpeg = "/9j/4QBDaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLwA8eDp4bXBtZXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEiLz7/4QAiRXhpZgAASUkqAAgAAAABABIBAwABAAAABgAAAAAAAAD/4AAQSkZJRgABAgAAAQABAAD/wAARCAABAAIDAREAAhEBAxEB/9sAQwADAgIDAgIDAwMDBAMDBAUIBQUEBAUKBwcGCAwKDAwLCgsLDQ4SEA0OEQ4LCxAWEBETFBUVFQwPFxgWFBgSFBUU/9sAQwEDBAQFBAUJBQUJFA0LDRQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQU/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD4H8Q/8h/Uv+vmX/0M1/o1wJ/ySWU/9g1D/wBNRMOM/wDkp8z/AOv9b/05I//Z";
        const oriented = await processImageInput({
          type: "image",
          data: orientedJpeg,
          mimeType: "image/jpeg",
        }, {
          resizeOptions: {
            maxWidth: 100,
            maxHeight: 100,
            maxBytes: orientedJpeg.length,
            jpegQuality: 80,
          },
        });

        const originalCreateElement = document.createElement.bind(document);
        document.createElement = (tagName, options) => {
          if (String(tagName).toLowerCase() === "canvas") {
            throw new Error("main-thread canvas use is forbidden during image processing");
          }
          return originalCreateElement(tagName, options);
        };
        let offMainThread;
        try {
          offMainThread = await processImageInput(input, {
            resizeOptions: { maxWidth: 30, maxHeight: 30 },
          });
        } finally {
          document.createElement = originalCreateElement;
        }

        const largePng = canvasBase64(3000, 3000, (context, canvas) => {
          const gradient = context.createLinearGradient(0, 0, canvas.width, canvas.height);
          gradient.addColorStop(0, "#001122");
          gradient.addColorStop(0.5, "#cc3366");
          gradient.addColorStop(1, "#ffee88");
          context.fillStyle = gradient;
          context.fillRect(0, 0, canvas.width, canvas.height);
        });
        let heartbeats = 0;
        const heartbeat = setInterval(() => { heartbeats += 1; }, 1);
        const responsive = await processImageInput({
          type: "image",
          data: largePng,
          mimeType: "image/png",
        }, {
          resizeOptions: { maxWidth: 500, maxHeight: 500, maxBytes: 100000, jpegQuality: 80 },
        });
        clearInterval(heartbeat);

        const abortController = new AbortController();
        const abortStarted = performance.now();
        const abortPromise = processImageInput({
          type: "image",
          data: largePng,
          mimeType: "image/png",
        }, {
          resizeOptions: { maxWidth: 3000, maxHeight: 3000, maxBytes: 1, jpegQuality: 100 },
          signal: abortController.signal,
        });
        setTimeout(() => abortController.abort(), 0);
        let abortResult;
        try {
          await abortPromise;
          abortResult = { name: "resolved", elapsedMs: performance.now() - abortStarted };
        } catch (error) {
          abortResult = { name: error.name, elapsedMs: performance.now() - abortStarted };
        }

        return {
          resized: {
            dimensions: await imageDimensions(resized),
            changed: resized.data !== original,
            hintText: normalized.hints.join("\\n"),
          },
          disabled: {
            unchanged: disabled.image.data === original,
            hints: disabled.hints,
          },
          toolResult: {
            types: toolResult.map((block) => block.type),
            dimensions: await imageDimensions(toolResult[1]),
            hint: toolResult[2].text,
          },
          byteOnly: byteOnly.ok ? {
            mimeType: byteOnly.image.mimeType,
            dimensions: await imageDimensions(byteOnly.image),
            underCap: byteOnly.image.data.length < qualityCap,
            progressedPast85: qualityCap < jpeg85Size && qualityCap > jpeg70Size,
            hints: byteOnly.hints,
          } : byteOnly,
          invalid,
          converted: converted.ok ? {
            mimeType: converted.image.mimeType,
            dimensions: await imageDimensions(converted.image),
            hints: converted.hints,
          } : converted,
          convertedWithResize: convertedWithResize.ok ? {
            mimeType: convertedWithResize.image.mimeType,
            dimensions: await imageDimensions(convertedWithResize.image),
            hints: convertedWithResize.hints,
          } : convertedWithResize,
          oriented: oriented.ok ? {
            dimensions: await imageDimensions(oriented.image),
            hints: oriented.hints,
          } : oriented,
          offMainThread: {
            ok: offMainThread.ok,
            dimensions: offMainThread.ok ? await imageDimensions(offMainThread.image) : null,
          },
          responsive: {
            ok: responsive.ok,
            heartbeatObserved: heartbeats > 0,
          },
          abortResult: {
            name: abortResult.name,
            prompt: abortResult.elapsedMs < 1000,
          },
        };
      })()
    `);

    assert.deepEqual(result.resized, {
      dimensions: { width: 40, height: 27 },
      changed: true,
      hintText: "[Image: original 120x80, displayed at 40x27. Multiply coordinates by 3.00 to map to original image.]",
    });
    assert.deepEqual(result.disabled, { unchanged: true, hints: [] });
    assert.deepEqual(result.toolResult, {
      types: ["text", "image", "text"],
      dimensions: { width: 40, height: 27 },
      hint: "[Image: original 120x80, displayed at 40x27. Multiply coordinates by 3.00 to map to original image.]",
    });

    assert.deepEqual(result.byteOnly, {
      mimeType: "image/jpeg",
      dimensions: { width: 180, height: 120 },
      underCap: true,
      progressedPast85: true,
      hints: [
        "[Image: original 180x120, displayed at 180x120. Multiply coordinates by 1.00 to map to original image.]",
      ],
    });

    assert.deepEqual(result.invalid, {
      ok: false,
      message: "[Image omitted: could not be converted to a supported inline image format.]",
    });
    const convertedBmp = {
      mimeType: "image/png",
      dimensions: { width: 1, height: 1 },
      hints: ["[Image converted from image/bmp to image/png.]"],
    };
    assert.deepEqual(result.converted, convertedBmp);
    assert.deepEqual(result.convertedWithResize, convertedBmp);
    assert.deepEqual(result.oriented, {
      dimensions: { width: 1, height: 2 },
      hints: [
        "[Image: original 1x2, displayed at 1x2. Multiply coordinates by 1.00 to map to original image.]",
      ],
    });
    assert.deepEqual(result.offMainThread, {
      ok: true,
      dimensions: { width: 30, height: 20 },
    });
    assert.deepEqual(result.responsive, {
      ok: true,
      heartbeatObserved: true,
    });
    assert.deepEqual(result.abortResult, {
      name: "AbortError",
      prompt: true,
    });
  } finally {
    await opened.finish();
  }
});
