/**
 * Test-time ESM resolver for Node's --experimental-strip-types runner.
 *
 * Source files use explicit ".js" import specifiers (for browser/bundler output),
 * while tests execute TypeScript sources directly. This loader retries relative
 * ".js" specifiers as ".ts" so node --test can resolve source modules.
 *
 * It also supports Vite-style `?raw` text imports for local markdown files used
 * by bundled docs/skills catalogs.
 */

import { readFile } from "node:fs/promises";

function hasCode(error) {
  return typeof error === "object" && error !== null && "code" in error;
}

function toRawFileUrl(rawUrl) {
  const url = new URL(rawUrl);
  url.search = "";
  url.hash = "";
  return url;
}

export async function resolve(specifier, context, defaultResolve) {
  try {
    return await defaultResolve(specifier, context, defaultResolve);
  } catch (error) {
    const isModuleNotFound = hasCode(error) && error.code === "ERR_MODULE_NOT_FOUND";
    const isRelative = specifier.startsWith("./") || specifier.startsWith("../");

    if (!isModuleNotFound || !isRelative) {
      throw error;
    }

    const [pathPart, queryPart] = specifier.split("?", 2);
    if (!pathPart.endsWith(".js")) {
      throw error;
    }

    const suffix = queryPart ? `?${queryPart}` : "";
    const tsSpecifier = `${pathPart.slice(0, -3)}.ts${suffix}`;
    return defaultResolve(tsSpecifier, context, defaultResolve);
  }
}

export async function load(url, context, defaultLoad) {
  if (url.endsWith("?raw")) {
    const fileUrl = toRawFileUrl(url);
    const sourceText = await readFile(fileUrl, "utf8");
    const source = `export default ${JSON.stringify(sourceText)};`;
    return {
      format: "module",
      shortCircuit: true,
      source,
    };
  }

  return defaultLoad(url, context, defaultLoad);
}
