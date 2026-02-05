/**
 * Snake game extension for Pi for Excel.
 * Browser-based version using canvas overlay.
 */

import type { ExcelExtensionAPI } from "../commands/extension-api.js";

const CELL = 16;
const COLS = 20;
const ROWS = 15;
const TICK = 120;

type Dir = "up" | "down" | "left" | "right";
type Pt = { x: number; y: number };

export function activate(api: ExcelExtensionAPI) {
  api.registerCommand("snake", {
    description: "Play Snake! 🐍",
    handler: () => {
      const el = document.createElement("div");
      el.style.cssText = `
        display: flex; flex-direction: column; align-items: center; gap: 8px;
        background: oklch(1 0 0 / 0.92); border-radius: 16px; padding: 16px;
        backdrop-filter: blur(24px); box-shadow: 0 8px 32px oklch(0 0 0 / 0.12);
      `;

      const header = document.createElement("div");
      header.style.cssText = "font-family: var(--font-mono); font-size: 12px; color: var(--muted-foreground);";
      header.textContent = "Score: 0 · Arrow keys to move · ESC to quit";
      el.appendChild(header);

      const canvas = document.createElement("canvas");
      canvas.width = COLS * CELL;
      canvas.height = ROWS * CELL;
      canvas.style.cssText = `border-radius: 8px; border: 1px solid oklch(0 0 0 / 0.08);`;
      el.appendChild(canvas);

      const ctx = canvas.getContext("2d")!;
      let snake: Pt[] = [{ x: 10, y: 7 }, { x: 9, y: 7 }, { x: 8, y: 7 }];
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
        // Background
        ctx.fillStyle = "#f8f8f6";
        ctx.fillRect(0, 0, canvas.width, canvas.height);

        // Grid (subtle)
        ctx.strokeStyle = "rgba(0,0,0,0.03)";
        for (let x = 0; x <= COLS; x++) { ctx.beginPath(); ctx.moveTo(x * CELL, 0); ctx.lineTo(x * CELL, ROWS * CELL); ctx.stroke(); }
        for (let y = 0; y <= ROWS; y++) { ctx.beginPath(); ctx.moveTo(0, y * CELL); ctx.lineTo(COLS * CELL, y * CELL); ctx.stroke(); }

        // Food
        ctx.fillStyle = "#d44";
        ctx.beginPath();
        ctx.arc(food.x * CELL + CELL / 2, food.y * CELL + CELL / 2, CELL / 2.5, 0, Math.PI * 2);
        ctx.fill();

        // Snake
        snake.forEach((p, i) => {
          const r = i === 0 ? 4 : 3;
          ctx.fillStyle = i === 0 ? "oklch(0.40 0.12 160)" : "oklch(0.50 0.10 160)";
          ctx.beginPath();
          ctx.roundRect(p.x * CELL + 1, p.y * CELL + 1, CELL - 2, CELL - 2, r);
          ctx.fill();
        });

        if (gameOver) {
          ctx.fillStyle = "rgba(0,0,0,0.5)";
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          ctx.fillStyle = "white";
          ctx.font = "bold 18px sans-serif";
          ctx.textAlign = "center";
          ctx.fillText("GAME OVER", canvas.width / 2, canvas.height / 2 - 10);
          ctx.font = "13px sans-serif";
          ctx.fillText(`Score: ${score} · Press R to restart`, canvas.width / 2, canvas.height / 2 + 14);
        }
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
          header.textContent = `Game Over! Score: ${score} · R to restart · ESC to quit`;
          return;
        }

        snake.unshift(nh);
        if (nh.x === food.x && nh.y === food.y) {
          score += 10;
          food = spawnFood(snake);
          header.textContent = `Score: ${score} · Arrow keys to move · ESC to quit`;
        } else {
          snake.pop();
        }
        draw();
      }

      const keyHandler = (e: KeyboardEvent) => {
        if (e.key === "Escape") { cleanup(); api.overlay.dismiss(); return; }
        if (e.key === "r" || e.key === "R") {
          if (gameOver) {
            snake = [{ x: 10, y: 7 }, { x: 9, y: 7 }, { x: 8, y: 7 }];
            food = spawnFood(snake);
            dir = "right"; nextDir = "right"; score = 0; gameOver = false;
            header.textContent = "Score: 0 · Arrow keys to move · ESC to quit";
            draw();
          }
          return;
        }
        const map: Record<string, Dir> = { ArrowUp: "up", ArrowDown: "down", ArrowLeft: "left", ArrowRight: "right", w: "up", s: "down", a: "left", d: "right" };
        const d = map[e.key];
        if (d) {
          e.preventDefault();
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

      api.overlay.show(el);
    },
  });
}
