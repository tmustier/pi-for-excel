declare const self: {
  addEventListener(type: "message", listener: (event: MessageEvent) => void): void;
  postMessage(message: unknown): void;
};

export interface PyodideWorkerRequest {
  id: string;
  type: "run";
  code: string;
  inputJson?: string;
  packages?: string[];
}

export interface PyodideWorkerResponse {
  id: string;
  ok: boolean;
  stdout?: string;
  stderr?: string;
  resultJson?: string;
  error?: string;
  loadTimeMs?: number;
  runTimeMs?: number;
}

interface PyodideInterface {
  runPythonAsync(code: string): Promise<unknown>;
  loadPackage(names: string | string[]): Promise<void>;
  loadPackagesFromImports(code: string): Promise<void>;
  globals: {
    set(name: string, value: unknown): void;
    get(name: string): unknown;
    delete(name: string): void;
  };
  toPy(value: unknown): unknown;
  version: string;
  FS: {
    writeFile(path: string, data: string | Uint8Array): void;
    readFile(path: string, options?: { encoding: string }): string;
  };
}

declare function _loadPyodide(options?: {
  indexURL?: string;
  stdout?: (text: string) => void;
  stderr?: (text: string) => void;
}): Promise<PyodideInterface>;

const PYODIDE_CDN_URL = "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";

let pyodide: PyodideInterface | null = null;
let loadingPromise: Promise<PyodideInterface> | null = null;

async function ensurePyodide(): Promise<PyodideInterface> {
  if (pyodide) return pyodide;
  if (loadingPromise) return loadingPromise;

  loadingPromise = (async () => {
    try {
      const mod = await import(/* @vite-ignore */ `${PYODIDE_CDN_URL}pyodide.mjs`) as {
        loadPyodide: typeof _loadPyodide;
      };

      const instance = await mod.loadPyodide({
        indexURL: PYODIDE_CDN_URL,
      });

      pyodide = instance;
      return instance;
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      throw new Error(
        "Failed to load Pyodide from jsDelivr. " +
        "Check internet access and Content-Security-Policy allowlists for https://cdn.jsdelivr.net. " +
        `Original error: ${reason}`,
      );
    } finally {
      if (!pyodide) {
        loadingPromise = null;
      }
    }
  })();

  return loadingPromise;
}

function detectImports(code: string): string[] {
  const imports = new Set<string>();
  const lines = code.split("\n");

  for (const line of lines) {
    const trimmed = line.trim();
    const importMatch = /^import\s+([\w.]+)/.exec(trimmed);
    if (importMatch) {
      const importName = importMatch[1];
      if (importName) {
        const topLevel = importName.split(".")[0];
        if (topLevel) imports.add(topLevel);
      }
    }
    const fromMatch = /^from\s+([\w.]+)\s+import/.exec(trimmed);
    if (fromMatch) {
      const importName = fromMatch[1];
      if (importName) {
        const topLevel = importName.split(".")[0];
        if (topLevel) imports.add(topLevel);
      }
    }
  }

  const SKIP = new Set([
    "sys", "os", "io", "re", "json", "math", "csv", "datetime",
    "collections", "itertools", "functools", "operator", "string",
    "textwrap", "unicodedata", "struct", "copy", "pprint",
    "typing", "types", "abc", "enum", "dataclasses",
    "pathlib", "tempfile", "shutil", "glob",
    "hashlib", "hmac", "secrets", "base64", "binascii",
    "html", "xml", "email", "urllib", "http",
    "logging", "warnings", "traceback",
    "time", "calendar", "random", "statistics",
    "decimal", "fractions",
    "threading", "multiprocessing", "subprocess", "signal",
    "socket", "ssl", "select",
    "sqlite3", "dbm",
    "gzip", "bz2", "zipfile", "tarfile", "lzma", "zlib",
    "unittest", "doctest",
    "argparse", "configparser", "getopt",
    "ctypes", "array", "weakref", "gc",
    "inspect", "dis", "ast", "code", "codeop",
    "pickle", "shelve", "marshal",
    "platform", "sysconfig", "site",
    // Pyodide distribution packages (pre-compiled WASM): not preloaded, but
    // loadPackagesFromImports loads them, so micropip must not try PyPI for them.
    "numpy", "np",
    "pandas", "pd",
    "scipy",
    "matplotlib", "mpl", "plt",
    "seaborn", "sns",
    "sklearn", "scikit_learn",
    "PIL", "Pillow",
    "lxml",
    "bs4", "beautifulsoup4",
    "pyodide", "micropip",
    "js",
  ]);

  return [...imports].filter((pkg) => !SKIP.has(pkg));
}

async function handleRun(request: PyodideWorkerRequest): Promise<PyodideWorkerResponse> {
  const loadStart = performance.now();
  const py = await ensurePyodide();
  const loadTimeMs = Math.round(performance.now() - loadStart);

  let stdout = "";
  let stderr = "";

  if (request.packages && request.packages.length > 0) {
    await py.loadPackage(request.packages);
  }

  try {
    await py.loadPackagesFromImports(request.code);
  } catch {
  }

  const detectedPackages = detectImports(request.code);
  if (detectedPackages.length > 0) {
    try {
      await py.loadPackage("micropip");
      const packageList = detectedPackages.map((p) => `"${p}"`).join(", ");
      await py.runPythonAsync(`
import micropip
try:
    await micropip.install([${packageList}])
except Exception:
    pass
`);
    } catch {
    }
  }

  if (request.inputJson) {
    py.globals.set("__input_json__", request.inputJson);
    await py.runPythonAsync(`
import json as __json__
input_data = __json__.loads(__input_json__)
del __input_json__
`);
  } else {
    await py.runPythonAsync("input_data = None");
  }

  await py.runPythonAsync(`
import sys as __sys__
import io as __io__
__stdout_capture__ = __io__.StringIO()
__stderr_capture__ = __io__.StringIO()
__sys__.stdout = __stdout_capture__
__sys__.stderr = __stderr_capture__
result = None
`);

  const runStart = performance.now();

  try {
    await py.runPythonAsync(request.code);
  } catch (error) {
    await py.runPythonAsync(`
__sys__.stdout = __sys__.__stdout__
__sys__.stderr = __sys__.__stderr__
`).catch(() => {});

    try {
      const getVal = await py.runPythonAsync("__stdout_capture__.getvalue()");
      stdout = getVal != null ? `${getVal as string}` : "";
    } catch {}

    try {
      const getVal = await py.runPythonAsync("__stderr_capture__.getvalue()");
      stderr = getVal != null ? `${getVal as string}` : "";
    } catch {}

    const message = error instanceof Error ? error.message : String(error);
    const response: PyodideWorkerResponse = {
      id: request.id,
      ok: false,
      error: message,
      loadTimeMs,
      runTimeMs: Math.round(performance.now() - runStart),
    };
    if (stdout.length > 0) response.stdout = stdout;
    if (stderr.length > 0) response.stderr = stderr;
    return response;
  }

  const runTimeMs = Math.round(performance.now() - runStart);

  try {
    const getStdout = await py.runPythonAsync("__stdout_capture__.getvalue()");
    stdout = getStdout != null ? `${getStdout as string}` : "";
  } catch {}

  try {
    const getStderr = await py.runPythonAsync("__stderr_capture__.getvalue()");
    stderr = getStderr != null ? `${getStderr as string}` : "";
  } catch {}

  let resultJson: string | undefined;
  try {
    const resultVal = await py.runPythonAsync(`
import json as __json__
__result_json__ = __json__.dumps(result) if result is not None else None
__result_json__
`);
    if (resultVal != null) {
      resultJson = `${resultVal as string}`;
    }
  } catch {
  }

  await py.runPythonAsync(`
__sys__.stdout = __sys__.__stdout__
__sys__.stderr = __sys__.__stderr__
`).catch(() => {});

  try {
    py.globals.delete("__stdout_capture__");
    py.globals.delete("__stderr_capture__");
    py.globals.delete("__result_json__");
  } catch {}

  const response: PyodideWorkerResponse = {
    id: request.id,
    ok: true,
    loadTimeMs,
    runTimeMs,
  };
  if (stdout.length > 0) response.stdout = stdout;
  if (stderr.length > 0) response.stderr = stderr;
  if (resultJson !== undefined) response.resultJson = resultJson;
  return response;
}

self.addEventListener("message", (event: MessageEvent<PyodideWorkerRequest>) => {
  const request = event.data;

  void handleRun(request).then(
    (response) => self.postMessage(response),
    (error: unknown) => {
      const message = error instanceof Error ? error.message : String(error);
      const response: PyodideWorkerResponse = {
        id: request.id,
        ok: false,
        error: `Worker error: ${message}`,
      };
      self.postMessage(response);
    },
  );
});
