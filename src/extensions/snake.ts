/**
 * Snake game extension for Pi for Excel.
 * Renders as an inline widget above the input — messages stay visible above.
 */

import type { ExcelExtensionAPI } from "../commands/extension-api.js";

const CELL = 14;
const COLS = 22;
const ROWS = 12;
const TICK = 110;

type Dir = "up" | "down" | "left" | "right";
type Pt = { x: number; y: number };

export function activate(api: ExcelExtensionAPI) {
  api.registerCommand("snake", {
    description: "Play Snake! 🐍",
    busyAllowed: true,
    handler: () => {
      const el = document.createElement("div");
      el.style.cssText = `
        display: flex; flex-direction: column; align-items: center; gap: 6px;
        background: oklch(1 0 0 / 0.92); border-radius: 12px; padding: 10px;
        backdrop-filter: blur(24px); box-shadow: 0 4px 16px oklch(0 0 0 / 0.08);
      `;

      const header = document.createElement("div");
      header.style.cssText = "font-family: var(--font-mono); font-size: 10.5px; color: var(--muted-foreground); width: 100%; display: flex; justify-content: space-between;";
      header.innerHTML = `<span>Score: 0</span><span style="opacity: 0.5">arrows · esc quit</span>`;
      el.appendChild(header);

      const canvas = document.createElement("canvas");
      const dpr = window.devicePixelRatio || 1;
      const logicalW = COLS * CELL;
      const logicalH = ROWS * CELL;
      canvas.width = logicalW * dpr;
      canvas.height = logicalH * dpr;
      canvas.style.cssText = `border-radius: 6px; border: 1px solid oklch(0 0 0 / 0.06); display: block; width: ${logicalW}px; height: ${logicalH}px;`;
      el.appendChild(canvas);

      const ctx = canvas.getContext("2d");
      if (!ctx) {
        throw new Error("Canvas 2D context not supported");
      }
      const ctx2d = ctx;
      ctx2d.scale(dpr, dpr);
      let snake: Pt[] = [{ x: 11, y: 6 }, { x: 10, y: 6 }, { x: 9, y: 6 }];
      let food = spawnFood(snake);
      let dir: Dir = "right";
      let nextDir: Dir = "right";
      let score = 0;
      let gameOver = false;

      function spawnFood(s: Pt[]): Pt {
        let f: Pt;
        do { f = { x: Math.floor(Math.random() * COLS), y: Math.floor(Math.random() * ROWS) }; }
        while (s.some(p => p.x === f.x && p.y === f.y));
        return f;
      }

      function draw() {
        ctx2d.fillStyle = "#f8f8f6";
        ctx2d.fillRect(0, 0, logicalW, logicalH);

        // Grid (subtle)
        ctx2d.strokeStyle = "rgba(0,0,0,0.03)";
        for (let x = 0; x <= COLS; x++) { ctx2d.beginPath(); ctx2d.moveTo(x * CELL, 0); ctx2d.lineTo(x * CELL, ROWS * CELL); ctx2d.stroke(); }
        for (let y = 0; y <= ROWS; y++) { ctx2d.beginPath(); ctx2d.moveTo(0, y * CELL); ctx2d.lineTo(COLS * CELL, y * CELL); ctx2d.stroke(); }

        // Food
        ctx2d.fillStyle = "#d44";
        ctx2d.beginPath();
        ctx2d.arc(food.x * CELL + CELL / 2, food.y * CELL + CELL / 2, CELL / 2.5, 0, Math.PI * 2);
        ctx2d.fill();

        // Snake
        snake.forEach((p, i) => {
          const r = i === 0 ? 3 : 2;
          ctx2d.fillStyle = i === 0 ? "oklch(0.40 0.12 160)" : "oklch(0.50 0.10 160)";
          ctx2d.beginPath();
          ctx2d.roundRect(p.x * CELL + 1, p.y * CELL + 1, CELL - 2, CELL - 2, r);
          ctx2d.fill();
        });

        if (gameOver) {
          ctx2d.fillStyle = "rgba(0,0,0,0.5)";
          ctx2d.fillRect(0, 0, logicalW, logicalH);
          ctx2d.fillStyle = "white";
          ctx2d.font = "bold 16px sans-serif";
          ctx2d.textAlign = "center";
          ctx2d.fillText("GAME OVER", logicalW / 2, logicalH / 2 - 8);
          ctx2d.font = "12px sans-serif";
          ctx2d.fillText(`Score: ${score} · R restart · ESC quit`, logicalW / 2, logicalH / 2 + 12);
        }
      }

      function updateHeader() {
        header.innerHTML = gameOver
          ? `<span>Game Over! ${score}</span><span style="opacity: 0.5">R restart · esc quit</span>`
          : `<span>Score: ${score}</span><span style="opacity: 0.5">arrows · esc quit</span>`;
      }

      function tick() {
        if (gameOver) return;
        dir = nextDir;
        const head = snake[0];
        const moves: Record<Dir, Pt> = {
          up: { x: head.x, y: head.y - 1 }, down: { x: head.x, y: head.y + 1 },
          left: { x: head.x - 1, y: head.y }, right: { x: head.x + 1, y: head.y },
        };
        const nh = moves[dir];

        if (nh.x < 0 || nh.x >= COLS || nh.y < 0 || nh.y >= ROWS || snake.some(s => s.x === nh.x && s.y === nh.y)) {
          gameOver = true;
          draw();
          updateHeader();
          return;
        }

        snake.unshift(nh);
        if (nh.x === food.x && nh.y === food.y) {
          score += 10;
          food = spawnFood(snake);
          updateHeader();
        } else {
          snake.pop();
        }
        draw();
      }

      const keyHandler = (e: KeyboardEvent) => {
        if (e.key === "Escape") { e.stopPropagation(); e.preventDefault(); cleanup(); api.widget.dismiss(); return; }
        if (e.key === "r" || e.key === "R") {
          if (gameOver) {
            snake = [{ x: 11, y: 6 }, { x: 10, y: 6 }, { x: 9, y: 6 }];
            food = spawnFood(snake);
            dir = "right"; nextDir = "right"; score = 0; gameOver = false;
            updateHeader();
            draw();
          }
          return;
        }
        const map: Record<string, Dir> = { ArrowUp: "up", ArrowDown: "down", ArrowLeft: "left", ArrowRight: "right", w: "up", s: "down", a: "left", d: "right" };
        const d = map[e.key];
        if (d) {
          e.preventDefault();
          e.stopPropagation();
          const opp: Record<Dir, Dir> = { up: "down", down: "up", left: "right", right: "left" };
          if (d !== opp[dir]) nextDir = d;
        }
      };

      document.addEventListener("keydown", keyHandler);
      const interval = setInterval(tick, TICK);
      draw();

      function cleanup() {
        clearInterval(interval);
        document.removeEventListener("keydown", keyHandler);
      }

      api.widget.show(el);
    },
  });
}
