#!/usr/bin/env node

/**
 * Local Python / LibreOffice bridge for Pi for Excel.
 *
 * Modes:
 * - stub (default): deterministic simulated responses for local development.
 * - real: executes local python + libreoffice commands with guardrails.
 *
 * Endpoints:
 * - GET  /health
 * - POST /v1/python-run
 * - POST /v1/libreoffice-convert
 */

import http from "node:http";
import https from "node:https";
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { spawn, spawnSync } from "node:child_process";
import { timingSafeEqual } from "node:crypto";

const args = new Set(process.argv.slice(2));
const useHttps = args.has("--https") || process.env.HTTPS === "1" || process.env.HTTPS === "true";
const useHttp = args.has("--http");

if (useHttps && useHttp) {
  console.error("[pi-for-excel] Invalid args: can't use both --https and --http");
  process.exit(1);
}

const HOST = process.env.HOST || (useHttps ? "localhost" : "127.0.0.1");
const PORT = Number.parseInt(process.env.PORT || "3340", 10);

const MODE_RAW = (process.env.PYTHON_BRIDGE_MODE || "stub").trim().toLowerCase();
const MODE = MODE_RAW === "real" ? "real" : MODE_RAW === "stub" ? "stub" : null;
if (!MODE) {
  console.error(`[pi-for-excel] Invalid PYTHON_BRIDGE_MODE: ${MODE_RAW}. Use "stub" or "real".`);
  process.exit(1);
}

const PYTHON_BIN = (process.env.PYTHON_BRIDGE_PYTHON_BIN || "python3").trim();
const LIBREOFFICE_BIN_RAW = (process.env.PYTHON_BRIDGE_LIBREOFFICE_BIN || "").trim();
const LIBREOFFICE_CANDIDATES = LIBREOFFICE_BIN_RAW.length > 0
  ? [LIBREOFFICE_BIN_RAW]
  : ["soffice", "libreoffice"];

const rootDir = path.resolve(process.cwd());
const keyPath = path.join(rootDir, "key.pem");
const certPath = path.join(rootDir, "cert.pem");

const DEFAULT_ALLOWED_ORIGINS = new Set([
  "https://localhost:3000",
  "https://pi-for-excel.vercel.app",
]);

const MAX_JSON_BODY_BYTES = 512 * 1024;
const MAX_CODE_LENGTH = 40_000;
const MAX_INPUT_JSON_LENGTH = 200_000;
const MAX_OUTPUT_BYTES = 256 * 1024;

const PYTHON_DEFAULT_TIMEOUT_MS = 10_000;
const PYTHON_MIN_TIMEOUT_MS = 100;
const PYTHON_MAX_TIMEOUT_MS = 120_000;

const LIBREOFFICE_DEFAULT_TIMEOUT_MS = 60_000;
const LIBREOFFICE_MIN_TIMEOUT_MS = 1_000;
const LIBREOFFICE_MAX_TIMEOUT_MS = 300_000;

const LIBREOFFICE_TARGET_FORMATS = new Set(["csv", "pdf", "xlsx"]);
const RESULT_JSON_MARKER = "__PI_FOR_EXCEL_RESULT_JSON_V1__";

const allowedOrigins = (() => {
  const raw = process.env.ALLOWED_ORIGINS;
  if (!raw) return DEFAULT_ALLOWED_ORIGINS;

  const custom = new Set(
    raw
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean),
  );

  return custom.size > 0 ? custom : DEFAULT_ALLOWED_ORIGINS;
})();

const authToken = (() => {
  const raw = process.env.PYTHON_BRIDGE_TOKEN;
  if (typeof raw !== "string") return "";
  return raw.trim();
})();

const PYTHON_WRAPPER_CODE = [
  "import json",
  "import os",
  "import sys",
  "import traceback",
  "",
  "raw_input = os.environ.get('PI_INPUT_JSON', '')",
  "input_data = None",
  "if raw_input:",
  "    input_data = json.loads(raw_input)",
  "",
  "user_code = os.environ.get('PI_USER_CODE', '')",
  "scope = {'__name__': '__main__', 'input_data': input_data}",
  "",
  "try:",
  "    exec(user_code, scope, scope)",
  "except Exception:",
  "    traceback.print_exc()",
  "    raise",
  "",
  "if 'result' in scope:",
  "    try:",
  `        print('${RESULT_JSON_MARKER}')`,
  "        print(json.dumps(scope['result'], ensure_ascii=False))",
  "    except Exception as exc:",
  "        print(f'[pi-python-bridge] result serialization error: {exc}', file=sys.stderr)",
].join("\n");

class HttpError extends Error {
  /**
   * @param {number} status
   * @param {string} message
   */
  constructor(status, message) {
    super(message);
    this.name = "HttpError";
    this.status = status;
  }
}

/**
 * @param {unknown} value
 * @returns {value is Record<string, unknown>}
 */
function isRecord(value) {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * @param {string | undefined} origin
 */
function isAllowedOrigin(origin) {
  return typeof origin === "string" && allowedOrigins.has(origin);
}

/**
 * @param {string | undefined} addr
 */
function isLoopbackAddress(addr) {
  if (!addr) return false;
  if (addr === "::1" || addr === "0:0:0:0:0:0:0:1") return true;
  if (addr.startsWith("127.")) return true;
  if (addr.startsWith("::ffff:127.")) return true;
  return false;
}

/**
 * @param {http.IncomingMessage} req
 * @param {http.ServerResponse} res
 */
function setCorsHeaders(req, res) {
  const origin = req.headers.origin;
  if (isAllowedOrigin(origin)) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Vary", "Origin");
  }

  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    req.headers["access-control-request-headers"] || "content-type,authorization",
  );
  res.setHeader("Access-Control-Max-Age", "86400");
}

/**
 * @param {http.ServerResponse} res
 * @param {number} status
 * @param {unknown} payload
 */
function respondJson(res, status, payload) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(payload));
}

/**
 * @param {http.ServerResponse} res
 * @param {number} status
 * @param {string} text
 */
function respondText(res, status, text) {
  res.statusCode = status;
  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.end(text);
}

/**
 * @param {string | undefined} headerValue
 */
function extractBearerToken(headerValue) {
  if (typeof headerValue !== "string") return null;
  const prefix = "Bearer ";
  if (!headerValue.startsWith(prefix)) return null;
  const token = headerValue.slice(prefix.length).trim();
  return token.length > 0 ? token : null;
}

/**
 * @param {string} left
 * @param {string} right
 */
function secureEquals(left, right) {
  const leftBuffer = Buffer.from(left);
  const rightBuffer = Buffer.from(right);
  if (leftBuffer.length !== rightBuffer.length) return false;
  return timingSafeEqual(leftBuffer, rightBuffer);
}

/**
 * @param {http.IncomingMessage} req
 */
function isAuthorized(req) {
  if (!authToken) return true;

  const candidate = extractBearerToken(req.headers.authorization);
  if (!candidate) return false;

  return secureEquals(candidate, authToken);
}

/**
 * @param {http.IncomingMessage} req
 */
async function readJsonBody(req) {
  const chunks = [];
  let size = 0;

  for await (const chunk of req) {
    const part = typeof chunk === "string" ? Buffer.from(chunk) : chunk;
    size += part.length;

    if (size > MAX_JSON_BODY_BYTES) {
      throw new HttpError(413, `Request body too large (max ${MAX_JSON_BODY_BYTES} bytes).`);
    }

    chunks.push(part);
  }

  const text = Buffer.concat(chunks).toString("utf8").trim();
  if (text.length === 0) {
    throw new HttpError(400, "Missing JSON request body.");
  }

  try {
    return JSON.parse(text);
  } catch {
    throw new HttpError(400, "Invalid JSON body.");
  }
}

/**
 * @param {unknown} value
 */
function normalizeOptionalString(value) {
  if (typeof value !== "string") return undefined;

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

/**
 * @param {unknown} value
 * @param {{ name: string; min: number; max: number; defaultValue: number }} options
 */
function parseBoundedInteger(value, options) {
  if (value === undefined) return options.defaultValue;
  if (typeof value !== "number" || !Number.isInteger(value)) {
    throw new HttpError(400, `${options.name} must be an integer.`);
  }
  if (value < options.min || value > options.max) {
    throw new HttpError(400, `${options.name} must be between ${options.min} and ${options.max}.`);
  }
  return value;
}

/**
 * @param {string} value
 */
function isAbsolutePath(value) {
  if (value.startsWith("/")) return true;
  if (/^[A-Za-z]:[\\/]/.test(value)) return true;
  if (value.startsWith("\\\\")) return true;
  return false;
}

/**
 * @param {unknown} value
 * @param {string} field
 */
function normalizeAbsolutePath(value, field) {
  const pathValue = normalizeOptionalString(value);
  if (!pathValue) {
    throw new HttpError(400, `${field} is required.`);
  }

  if (!isAbsolutePath(pathValue)) {
    throw new HttpError(400, `${field} must be an absolute path.`);
  }

  return pathValue;
}

/**
 * @param {string | undefined} output
 */
function normalizeOutput(output) {
  if (typeof output !== "string") return undefined;
  const normalized = output.replace(/\r/g, "").trim();
  return normalized.length > 0 ? normalized : undefined;
}

/**
 * @param {unknown} payload
 */
function parsePythonRunRequest(payload) {
  if (!isRecord(payload)) {
    throw new HttpError(400, "Request body must be a JSON object.");
  }

  const code = typeof payload.code === "string" ? payload.code : "";
  if (code.trim().length === 0) {
    throw new HttpError(400, "code is required.");
  }

  if (code.length > MAX_CODE_LENGTH) {
    throw new HttpError(400, `code is too long (max ${MAX_CODE_LENGTH} characters).`);
  }

  const inputJson = normalizeOptionalString(payload.input_json);
  if (inputJson && inputJson.length > MAX_INPUT_JSON_LENGTH) {
    throw new HttpError(400, `input_json is too long (max ${MAX_INPUT_JSON_LENGTH} characters).`);
  }

  if (inputJson) {
    try {
      void JSON.parse(inputJson);
    } catch {
      throw new HttpError(400, "input_json must be valid JSON.");
    }
  }

  const timeoutMs = parseBoundedInteger(payload.timeout_ms, {
    name: "timeout_ms",
    min: PYTHON_MIN_TIMEOUT_MS,
    max: PYTHON_MAX_TIMEOUT_MS,
    defaultValue: PYTHON_DEFAULT_TIMEOUT_MS,
  });

  return {
    code,
    input_json: inputJson,
    timeout_ms: timeoutMs,
  };
}

/**
 * @param {unknown} payload
 */
function parseLibreOfficeRequest(payload) {
  if (!isRecord(payload)) {
    throw new HttpError(400, "Request body must be a JSON object.");
  }

  const inputPath = normalizeAbsolutePath(payload.input_path, "input_path");

  const targetFormat = normalizeOptionalString(payload.target_format)?.toLowerCase();
  if (!targetFormat || !LIBREOFFICE_TARGET_FORMATS.has(targetFormat)) {
    throw new HttpError(400, "target_format must be one of: csv, pdf, xlsx.");
  }

  let outputPath;
  if (payload.output_path !== undefined) {
    outputPath = normalizeAbsolutePath(payload.output_path, "output_path");
  }

  const overwrite = typeof payload.overwrite === "boolean" ? payload.overwrite : false;

  const timeoutMs = parseBoundedInteger(payload.timeout_ms, {
    name: "timeout_ms",
    min: LIBREOFFICE_MIN_TIMEOUT_MS,
    max: LIBREOFFICE_MAX_TIMEOUT_MS,
    defaultValue: LIBREOFFICE_DEFAULT_TIMEOUT_MS,
  });

  return {
    input_path: inputPath,
    target_format: targetFormat,
    output_path: outputPath,
    overwrite,
    timeout_ms: timeoutMs,
  };
}

/**
 * @param {string} command
 * @param {string[]} args
 */
function probeBinary(command, args) {
  const probe = spawnSync(command, args, {
    encoding: "utf8",
  });

  if (probe.error) {
    console.error(`[pi-for-excel] Failed to probe binary "${command}":`, probe.error);
    return {
      available: false,
      error: "probe_failed",
      command,
    };
  }

  if (probe.status !== 0) {
    return {
      available: false,
      error: `probe_exit_${String(probe.status)}`,
      command,
    };
  }

  const stdout = typeof probe.stdout === "string" ? probe.stdout.trim() : "";
  return {
    available: true,
    version: stdout || command,
    command,
  };
}

function probeLibreOfficeBinary() {
  for (const candidate of LIBREOFFICE_CANDIDATES) {
    const probe = probeBinary(candidate, ["--version"]);
    if (probe.available) {
      return {
        available: true,
        command: candidate,
        version: probe.version,
      };
    }
  }

  return {
    available: false,
    command: LIBREOFFICE_CANDIDATES[0] || "soffice",
    error: `No LibreOffice binary found (tried: ${LIBREOFFICE_CANDIDATES.join(", ")})`,
  };
}

/**
 * @param {{ command: string; args: string[]; timeoutMs: number; env?: NodeJS.ProcessEnv }} options
 */
async function runCommandCapture(options) {
  let timedOut = false;
  let overflow = false;

  const result = await new Promise((resolve, reject) => {
    const child = spawn(options.command, options.args, {
      stdio: ["ignore", "pipe", "pipe"],
      env: options.env,
    });

    let stdout = "";
    let stderr = "";

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");

    child.stdout.on("data", (chunk) => {
      if (overflow) return;
      stdout += chunk;
      if (Buffer.byteLength(stdout, "utf8") > MAX_OUTPUT_BYTES) {
        overflow = true;
        child.kill("SIGKILL");
      }
    });

    child.stderr.on("data", (chunk) => {
      if (overflow) return;
      stderr += chunk;
      if (Buffer.byteLength(stderr, "utf8") > MAX_OUTPUT_BYTES) {
        overflow = true;
        child.kill("SIGKILL");
      }
    });

    const timeoutId = setTimeout(() => {
      timedOut = true;
      child.kill("SIGKILL");
    }, options.timeoutMs);

    child.once("error", (error) => {
      clearTimeout(timeoutId);
      reject(error);
    });

    child.once("close", (code, signal) => {
      clearTimeout(timeoutId);
      resolve({
        code: code ?? -1,
        signal: signal ?? null,
        stdout,
        stderr,
      });
    });
  });

  if (overflow) {
    throw new HttpError(413, `Process output exceeded ${MAX_OUTPUT_BYTES} bytes.`);
  }

  if (timedOut) {
    throw new HttpError(504, `Process timed out after ${options.timeoutMs}ms.`);
  }

  if (result.signal) {
    throw new HttpError(500, `Process exited with signal ${result.signal}.`);
  }

  return result;
}

/**
 * @param {string} stdout
 */
function splitPythonResult(stdout) {
  const normalized = stdout.replace(/\r/g, "");
  const marker = `\n${RESULT_JSON_MARKER}\n`;

  let markerIndex = normalized.lastIndexOf(marker);
  let markerLength = marker.length;

  if (markerIndex === -1 && normalized.startsWith(`${RESULT_JSON_MARKER}\n`)) {
    markerIndex = 0;
    markerLength = RESULT_JSON_MARKER.length + 1;
  }

  if (markerIndex === -1) {
    return {
      stdout: normalizeOutput(normalized),
      resultJson: undefined,
    };
  }

  const before = normalized.slice(0, markerIndex);
  const after = normalized.slice(markerIndex + markerLength).trim();

  return {
    stdout: normalizeOutput(before),
    resultJson: after.length > 0 ? after : undefined,
  };
}

/**
 * @param {{ code: string; input_json?: string; timeout_ms: number }} request
 * @param {{ command: string }} pythonInfo
 */
async function runPython(request, pythonInfo) {
  const env = {
    ...process.env,
    PI_USER_CODE: request.code,
    PI_INPUT_JSON: request.input_json || "",
  };

  const result = await runCommandCapture({
    command: pythonInfo.command,
    args: ["-I", "-c", PYTHON_WRAPPER_CODE],
    timeoutMs: request.timeout_ms,
    env,
  });

  if (result.code !== 0) {
    const message = [result.stderr, result.stdout]
      .map((value) => value.trim())
      .find((value) => value.length > 0) || `python exited with code ${result.code}`;

    throw new HttpError(400, message);
  }

  const parsed = splitPythonResult(result.stdout);

  if (parsed.resultJson) {
    try {
      void JSON.parse(parsed.resultJson);
    } catch {
      throw new HttpError(500, "Bridge produced invalid result_json payload.");
    }
  }

  return {
    ok: true,
    action: "run_python",
    exit_code: result.code,
    stdout: parsed.stdout,
    stderr: normalizeOutput(result.stderr),
    result_json: parsed.resultJson,
    truncated: false,
  };
}

/**
 * @param {{ input_path: string; target_format: string; output_path?: string; overwrite: boolean; timeout_ms: number }} request
 */
function resolveOutputPath(request) {
  if (request.output_path) {
    return request.output_path;
  }

  const parsed = path.parse(request.input_path);
  return path.join(parsed.dir, `${parsed.name}.${request.target_format}`);
}

/**
 * @param {string} filePath
 */
function ensureFileExists(filePath) {
  let stats;
  try {
    stats = fs.statSync(filePath);
  } catch {
    throw new HttpError(400, `input_path does not exist: ${filePath}`);
  }

  if (!stats.isFile()) {
    throw new HttpError(400, `input_path is not a file: ${filePath}`);
  }
}

/**
 * @param {string} outputPath
 * @param {boolean} overwrite
 */
function ensureOutputWritable(outputPath, overwrite) {
  if (fs.existsSync(outputPath) && !overwrite) {
    throw new HttpError(409, `output_path already exists: ${outputPath}`);
  }

  const outputDir = path.dirname(outputPath);
  if (!fs.existsSync(outputDir)) {
    throw new HttpError(400, `output_path directory does not exist: ${outputDir}`);
  }
}

/**
 * @param {string} tempDir
 * @param {string} inputPath
 * @param {string} targetFormat
 */
function findConvertedFile(tempDir, inputPath, targetFormat) {
  const baseName = path.parse(inputPath).name;
  const expected = path.join(tempDir, `${baseName}.${targetFormat}`);
  if (fs.existsSync(expected)) {
    return expected;
  }

  const entries = fs.readdirSync(tempDir, { withFileTypes: true });
  for (const entry of entries) {
    if (!entry.isFile()) continue;
    if (!entry.name.toLowerCase().endsWith(`.${targetFormat}`)) continue;
    return path.join(tempDir, entry.name);
  }

  return null;
}

/**
 * @param {{ input_path: string; target_format: string; output_path?: string; overwrite: boolean; timeout_ms: number }} request
 * @param {{ command: string }} libreOfficeInfo
 */
async function runLibreOfficeConvert(request, libreOfficeInfo) {
  ensureFileExists(request.input_path);

  const outputPath = resolveOutputPath(request);
  ensureOutputWritable(outputPath, request.overwrite);

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "pi-libreoffice-"));

  try {
    const result = await runCommandCapture({
      command: libreOfficeInfo.command,
      args: [
        "--headless",
        "--convert-to",
        request.target_format,
        "--outdir",
        tempDir,
        request.input_path,
      ],
      timeoutMs: request.timeout_ms,
      env: process.env,
    });

    if (result.code !== 0) {
      const message = [result.stderr, result.stdout]
        .map((value) => value.trim())
        .find((value) => value.length > 0) || `LibreOffice exited with code ${result.code}`;
      throw new HttpError(400, message);
    }

    const convertedPath = findConvertedFile(tempDir, request.input_path, request.target_format);
    if (!convertedPath) {
      throw new HttpError(500, "LibreOffice did not produce an output file.");
    }

    fs.copyFileSync(convertedPath, outputPath);

    const stats = fs.statSync(outputPath);

    return {
      ok: true,
      action: "convert",
      input_path: request.input_path,
      target_format: request.target_format,
      output_path: outputPath,
      bytes: stats.size,
      converter: libreOfficeInfo.command,
    };
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

function createStubBackend() {
  return {
    mode: "stub",
    async health() {
      return {
        backend: "stub",
        python: {
          available: true,
          mode: "stub",
        },
        libreoffice: {
          available: true,
          mode: "stub",
        },
      };
    },

    /**
     * @param {{ code: string; input_json?: string; timeout_ms: number }} request
     */
    async handlePython(request) {
      let resultJson;
      if (request.input_json) {
        resultJson = request.input_json;
      }

      return {
        ok: true,
        action: "run_python",
        exit_code: 0,
        stdout: "[stub] Python execution simulated.",
        result_json: resultJson,
        truncated: false,
      };
    },

    /**
     * @param {{ input_path: string; target_format: string; output_path?: string; overwrite: boolean; timeout_ms: number }} request
     */
    async handleLibreOffice(request) {
      const outputPath = resolveOutputPath(request);

      return {
        ok: true,
        action: "convert",
        input_path: request.input_path,
        target_format: request.target_format,
        output_path: outputPath,
        bytes: 0,
        converter: "stub",
      };
    },
  };
}

function createRealBackend() {
  const pythonInfo = probeBinary(PYTHON_BIN, ["--version"]);
  const libreOfficeInfo = probeLibreOfficeBinary();

  return {
    mode: "real",
    async health() {
      return {
        backend: "real",
        python: pythonInfo.available
          ? {
            available: true,
            command: pythonInfo.command,
            version: pythonInfo.version,
          }
          : {
            available: false,
            command: pythonInfo.command,
            error: pythonInfo.error,
          },
        libreoffice: libreOfficeInfo.available
          ? {
            available: true,
            command: libreOfficeInfo.command,
            version: libreOfficeInfo.version,
          }
          : {
            available: false,
            command: libreOfficeInfo.command,
            error: libreOfficeInfo.error,
          },
      };
    },

    /**
     * @param {{ code: string; input_json?: string; timeout_ms: number }} request
     */
    async handlePython(request) {
      if (!pythonInfo.available) {
        throw new HttpError(501, "python binary not available");
      }

      return runPython(request, { command: pythonInfo.command });
    },

    /**
     * @param {{ input_path: string; target_format: string; output_path?: string; overwrite: boolean; timeout_ms: number }} request
     */
    async handleLibreOffice(request) {
      if (!libreOfficeInfo.available) {
        throw new HttpError(501, "libreoffice binary not available");
      }

      return runLibreOfficeConvert(request, { command: libreOfficeInfo.command });
    },
  };
}

const backend = MODE === "real" ? createRealBackend() : createStubBackend();

const handler = async (req, res) => {
  try {
    const remote = req.socket?.remoteAddress;
    if (!isLoopbackAddress(remote)) {
      respondText(res, 403, "forbidden");
      return;
    }

    const origin = req.headers.origin;
    if (!isAllowedOrigin(origin)) {
      respondText(res, 403, "forbidden");
      return;
    }

    setCorsHeaders(req, res);

    if (req.method === "OPTIONS") {
      res.statusCode = 204;
      res.end();
      return;
    }

    const rawUrl = req.url || "/";
    const url = new URL(rawUrl, `http://${HOST}:${PORT}`);

    if (url.pathname === "/health") {
      if (req.method !== "GET") {
        throw new HttpError(405, "Method not allowed.");
      }

      respondJson(res, 200, {
        ok: true,
        mode: backend.mode,
        ...await backend.health(),
      });
      return;
    }

    if (url.pathname === "/v1/python-run") {
      if (req.method !== "POST") {
        throw new HttpError(405, "Method not allowed.");
      }

      if (!isAuthorized(req)) {
        throw new HttpError(401, "Unauthorized.");
      }

      const payload = await readJsonBody(req);
      const request = parsePythonRunRequest(payload);
      const result = await backend.handlePython(request);

      respondJson(res, 200, {
        ok: true,
        ...result,
      });
      return;
    }

    if (url.pathname === "/v1/libreoffice-convert") {
      if (req.method !== "POST") {
        throw new HttpError(405, "Method not allowed.");
      }

      if (!isAuthorized(req)) {
        throw new HttpError(401, "Unauthorized.");
      }

      const payload = await readJsonBody(req);
      const request = parseLibreOfficeRequest(payload);
      const result = await backend.handleLibreOffice(request);

      respondJson(res, 200, {
        ok: true,
        ...result,
      });
      return;
    }

    throw new HttpError(404, "Not found.");
  } catch (error) {
    const isHttpError = error instanceof HttpError;
    const status = isHttpError ? error.status : 500;

    if (!isHttpError) {
      console.error("[pi-for-excel] Unhandled python bridge error:", error);
    }

    const message = isHttpError
      ? error.message
      : "Internal server error.";

    respondJson(res, status, {
      ok: false,
      error: message,
    });
  }
};

const server = (() => {
  if (!useHttps) {
    return http.createServer(handler);
  }

  if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
    console.error("[pi-for-excel] HTTPS requested but key.pem/cert.pem not found in repo root.");
    console.error("Generate them with mkcert (see README). Example: mkcert localhost");
    process.exit(1);
  }

  return https.createServer(
    {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath),
    },
    handler,
  );
})();

server.listen(PORT, HOST, () => {
  const scheme = useHttps ? "https" : "http";
  console.log(`[pi-for-excel] python bridge listening on ${scheme}://${HOST}:${PORT}`);
  console.log(`[pi-for-excel] mode: ${backend.mode}`);
  console.log(`[pi-for-excel] health: ${scheme}://${HOST}:${PORT}/health`);
  console.log(`[pi-for-excel] endpoint: ${scheme}://${HOST}:${PORT}/v1/python-run`);
  console.log(`[pi-for-excel] endpoint: ${scheme}://${HOST}:${PORT}/v1/libreoffice-convert`);
  console.log(`[pi-for-excel] allowed origins: ${Array.from(allowedOrigins).join(", ")}`);

  if (authToken) {
    console.log("[pi-for-excel] auth: bearer token required for POST endpoints");
  }

  if (backend.mode === "stub") {
    console.log("[pi-for-excel] stub mode: python/libreoffice calls are simulated.");
    console.log("[pi-for-excel] use PYTHON_BRIDGE_MODE=real for local command execution.");
  }
});
