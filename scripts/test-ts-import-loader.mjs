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
