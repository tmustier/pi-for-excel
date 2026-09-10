import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import http from "node:http";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { test } from "node:test";

async function getFreePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.once("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      if (!address || typeof address === "string") {
        reject(new Error("Failed to reserve a port"));
        return;
      }
      server.close((error) => error ? reject(error) : resolve(address.port));
    });
  });
}

async function requestAuth(port: number, host: string): Promise<{ status: number; body: string; cacheControl?: string }> {
  return new Promise((resolve, reject) => {
    const request = http.get({ hostname: "127.0.0.1", port, path: "/__pi-auth", headers: { Host: host } }, (response) => {
      response.setEncoding("utf8");
      let body = "";
      response.on("data", (chunk: string) => { body += chunk; });
      response.on("end", () => resolve({
        status: response.statusCode ?? 0,
        body,
        ...(response.headers["cache-control"] ? { cacheControl: response.headers["cache-control"] } : {}),
      }));
    });
    request.once("error", reject);
  });
}

void test("dev auth endpoint serves credentials only for approved local HTTP requests", async (t) => {
  const home = await mkdtemp(path.join(os.tmpdir(), "pi-auth-http-"));
  const authDir = path.join(home, ".pi", "agent");
  await mkdir(authDir, { recursive: true });
  await writeFile(path.join(authDir, "auth.json"), '{"test":"credential"}');
  const port = await getFreePort();
  const child = spawn(path.resolve("node_modules/.bin/vite"), ["--host", "127.0.0.1", "--port", String(port), "--strictPort"], {
    cwd: path.resolve("."),
    env: { ...process.env, HOME: home },
    stdio: ["ignore", "pipe", "pipe"],
  });
  let output = "";
  child.stdout.setEncoding("utf8");
  child.stderr.setEncoding("utf8");
  child.stdout.on("data", (chunk: string) => { output += chunk; });
  child.stderr.on("data", (chunk: string) => { output += chunk; });
  t.after(async () => {
    child.kill("SIGTERM");
    await Promise.race([once(child, "exit"), delay(2_000)]).catch(() => undefined);
    await rm(home, { recursive: true, force: true });
  });

  let ready = false;
  for (let attempt = 0; attempt < 80; attempt += 1) {
    try {
      const response = await requestAuth(port, `localhost:${port}`);
      if (response.status === 200) {
        ready = true;
        break;
      }
    } catch {
      // Vite may still be starting.
    }
    await delay(50);
  }
  assert.equal(ready, true, `Vite did not start:\n${output}`);

  const cases = [
    { host: `localhost:${port}`, status: 200, body: /"credential"/, cacheControl: "no-store" },
    { host: `127.0.0.1:${port}`, status: 200, body: /"credential"/, cacheControl: "no-store" },
    { host: `[::1]:${port}`, status: 200, body: /"credential"/, cacheControl: "no-store" },
    { host: `10.0.2.2:${port}`, status: 403, body: /forbidden/, cacheControl: "no-store" },
    { host: `example.com:${port}`, status: 403, body: /not allowed/ },
  ];

  for (const entry of cases) {
    const response = await requestAuth(port, entry.host);
    assert.equal(response.status, entry.status, entry.host);
    assert.match(response.body, entry.body, entry.host);
    assert.equal(response.cacheControl, entry.cacheControl, entry.host);
  }
});
