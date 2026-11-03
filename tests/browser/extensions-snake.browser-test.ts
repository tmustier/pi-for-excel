import assert from "node:assert/strict";
import { after, before, test } from "node:test";
import { openBrowserPage, startTaskpaneServer, type TaskpaneServer } from "./harness.ts";

let env: TaskpaneServer;
before(async () => { env = await startTaskpaneServer(); });
after(async () => { await env.close(); });

void test("Snake fills the board without an unbounded food search", async () => {
  const { page, finish } = await openBrowserPage(env, { path: "src/ui-gallery.html" });
  try {
    const result = await page.evaluate(async () => {
      const snakePath = "/src/extensions/snake.ts";
      const apiPath = "/src/commands/extension-api.ts";
      const registryPath = "/src/commands/types.ts";
      const { activate } = await import(/* @vite-ignore */ snakePath) as typeof import("../../src/extensions/snake.ts");
      const { createExtensionAPI, loadExtension } = await import(/* @vite-ignore */ apiPath) as typeof import("../../src/commands/extension-api.ts");
      const { commandRegistry } = await import(/* @vite-ignore */ registryPath) as typeof import("../../src/commands/types.ts");
      const slot = document.getElementById("pi-widget-slot") ?? document.createElement("div");
      slot.id = "pi-widget-slot";
      document.body.appendChild(slot);
      const nativeSetInterval = window.setInterval;
      const nativeClearInterval = window.clearInterval;
      const nativeRandom = Math.random;
      let tick = (): void => {};
      let randomCalls = 0;
      Math.random = () => {
        randomCalls += 1;
        if (randomCalls > 1_000) throw new Error("food search exceeded its finite randomness budget");
        return 0;
      };
      window.setInterval = (handler: TimerHandler): number => {
        if (typeof handler !== "function") throw new Error("expected interval callback");
        tick = handler as () => void;
        return 1;
      };
      window.clearInterval = () => {};
      const api = createExtensionAPI({ getAgent: () => { throw new Error("unused"); }, widgetApiV2Enabled: false });
      const handle = await loadExtension(api, activate);
      let error = "";
      let steps = 0;
      try {
        await commandRegistry.get("snake")?.execute("");
        const cycle: Array<{ x: number; y: number }> = [];
        for (let x = 0; x < 22; x += 1) cycle.push({ x, y: 0 });
        for (let y = 1; y < 12; y += 1) {
          for (let offset = 0; offset < 21; offset += 1) cycle.push({ x: y % 2 === 0 ? offset + 1 : 21 - offset, y });
        }
        for (let y = 11; y > 0; y -= 1) cycle.push({ x: 0, y });
        let index = cycle.findIndex((point) => point.x === 11 && point.y === 6);
        while (!slot.textContent?.includes("Game Over!") && steps < 100_000) {
          const head = cycle[index];
          index = (index + 1) % cycle.length;
          const next = cycle[index];
          if (!head || !next) throw new Error("invalid cycle");
          const key = next.x > head.x ? "ArrowRight" : next.x < head.x ? "ArrowLeft" : next.y > head.y ? "ArrowDown" : "ArrowUp";
          document.dispatchEvent(new KeyboardEvent("keydown", { key, cancelable: true }));
          tick();
          steps += 1;
        }
      } catch (failure) { error = failure instanceof Error ? failure.message : String(failure); }
      const header = slot.textContent;
      await handle.deactivate();
      window.setInterval = nativeSetInterval;
      window.clearInterval = nativeClearInterval;
      Math.random = nativeRandom;
      return { error, header, steps };
    });
    assert.equal(result.error, "", "food search must terminate on a full board");
    assert.match(result.header ?? "", /Game Over! 2610/);
    assert.ok(result.steps < 100_000);
  } finally { await finish(); }
});

void test("Snake releases its interval and keyboard handler on extension teardown", async () => {
  const { page, finish } = await openBrowserPage(env, { path: "src/ui-gallery.html" });
  try {
    const result = await page.evaluate(async () => {
      const snakePath = "/src/extensions/snake.ts";
      const apiPath = "/src/commands/extension-api.ts";
      const registryPath = "/src/commands/types.ts";
      const { activate } = await import(/* @vite-ignore */ snakePath) as typeof import("../../src/extensions/snake.ts");
      const { createExtensionAPI, loadExtension } = await import(/* @vite-ignore */ apiPath) as typeof import("../../src/commands/extension-api.ts");
      const { commandRegistry } = await import(/* @vite-ignore */ registryPath) as typeof import("../../src/commands/types.ts");
      const slot = document.getElementById("pi-widget-slot") ?? document.createElement("div");
      slot.id = "pi-widget-slot";
      document.body.appendChild(slot);
      const nativeSetInterval = window.setInterval;
      const nativeClearInterval = window.clearInterval;
      const intervals = new Set<number>();
      let nextInterval = 1;
      window.setInterval = (): number => { const id = nextInterval++; intervals.add(id); return id; };
      window.clearInterval = (id?: number): void => { if (id !== undefined) intervals.delete(id); };
      const api = createExtensionAPI({ getAgent: () => { throw new Error("unused"); }, widgetApiV2Enabled: false });
      const handle = await loadExtension(api, activate);
      await commandRegistry.get("snake")?.execute("");
      const before = intervals.size;
      await handle.deactivate();
      const after = intervals.size;
      const key = new KeyboardEvent("keydown", { key: "ArrowUp", cancelable: true });
      document.dispatchEvent(key);
      window.setInterval = nativeSetInterval;
      window.clearInterval = nativeClearInterval;
      return { before, after, prevented: key.defaultPrevented };
    });
    assert.equal(result.before, 1);
    assert.equal(result.after, 0, "extension teardown must release the game interval");
    assert.equal(result.prevented, false, "extension teardown must release the game keyboard handler");
  } finally { await finish(); }
});
