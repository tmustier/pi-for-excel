import assert from "node:assert/strict";
import { after, before, test } from "node:test";

import { openTaskpane, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;

before(async () => {
  env = await startTaskpaneServer();
});

after(async () => {
  await env.close();
});

void test("browser image ingress resizes to the active model profile", async () => {
  const opened = await openTaskpane(env, { clientId: "image-limits-browser-client" });
  try {
    const result = await opened.page.evaluate(`
      (async () => {
        const { normalizePromptImages } = await import("/src/messages/image-input-limits.ts");
        const canvas = document.createElement("canvas");
        canvas.width = 120;
        canvas.height = 80;
        const context = canvas.getContext("2d");
        context.fillStyle = "#c2185b";
        context.fillRect(0, 0, canvas.width, canvas.height);
        const original = canvas.toDataURL("image/png").split(",")[1];
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
        const normalized = await normalizePromptImages([input], model);
        const disabled = await normalizePromptImages([input], model, { autoResizeImages: false });
        const resized = normalized.images[0];
        const image = new Image();
        const dimensions = await new Promise((resolve, reject) => {
          image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight });
          image.onerror = () => reject(new Error("resized image did not decode"));
          image.src = "data:" + resized.mimeType + ";base64," + resized.data;
        });
        return {
          dimensions,
          changed: resized.data !== original,
          hintText: normalized.hints.join("\\n"),
          disabled: {
            unchanged: disabled.images[0].data === original,
            hints: disabled.hints,
          },
        };
      })()
    `);

    assert.deepEqual(result, {
      dimensions: { width: 40, height: 27 },
      changed: true,
      hintText: "[Image: original 120x80, displayed at 40x27. Multiply coordinates by 3.00 to map to original image.]",
      disabled: { unchanged: true, hints: [] },
    });
  } finally {
    await opened.finish();
  }
});
