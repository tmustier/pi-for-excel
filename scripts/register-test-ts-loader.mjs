import { register } from "node:module";
import { pathToFileURL } from "node:url";

import { createRawMarkdownGlobLoader } from "./test-raw-markdown-glob.mjs";

Reflect.set(globalThis, "__PI_TEST_RAW_MARKDOWN_GLOB", createRawMarkdownGlobLoader());

register("./scripts/test-ts-import-loader.mjs", pathToFileURL("./"));
