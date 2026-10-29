import { marked, Renderer, type MarkedExtension, type Tokens } from "marked";

import {
  createMarkdownImageRenderPlan,
  isAllowedMarkdownUrl,
  isMarkdownExtensionDisabledByPolicy,
} from "./marked-safety-policy.js";
import { escapeAttr, escapeHtml } from "../utils/html.js";

let installed = false;

const unpatchedLink: (this: Renderer, token: Tokens.Link) => string = Renderer.prototype.link;

function safeLink(this: Renderer, token: Tokens.Link): string {
  const href = typeof token.href === "string" ? token.href : "";

  if (!isAllowedMarkdownUrl(href)) {
    const text = typeof token.text === "string" && token.text.trim().length > 0
      ? token.text
      : href;
    return escapeHtml(text);
  }

  return unpatchedLink.call(this, token);
}

function safeImage(token: Tokens.Image): string {
  const href = typeof token.href === "string" ? token.href : "";
  const alt = typeof token.text === "string" ? token.text : "";

  const plan = createMarkdownImageRenderPlan(href, alt);
  if (plan.kind === "link") {
    return `<a href="${escapeAttr(plan.href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(plan.label)}</a>`;
  }

  return escapeHtml(plan.label);
}

export function createSafeRenderer(): Renderer {
  const renderer = new Renderer();
  renderer.link = safeLink;
  renderer.image = safeImage;
  return renderer;
}

function withoutDisabledMarkedExtensions(extension: MarkedExtension): MarkedExtension {
  const extensionList = extension.extensions;
  if (!Array.isArray(extensionList) || extensionList.length === 0) {
    return extension;
  }

  const filteredExtensions = extensionList.filter(
    (definition) => !isMarkdownExtensionDisabledByPolicy(definition.name),
  );

  if (filteredExtensions.length === extensionList.length) {
    return extension;
  }

  return {
    ...extension,
    extensions: filteredExtensions,
  };
}

export function installMarkedSafetyPatch(): void {
  if (installed) return;
  installed = true;

  const originalUse = marked.use;
  marked.use = function patchedMarkedUse(
    this: typeof marked,
    ...extensions: MarkedExtension[]
  ): typeof marked {
    return originalUse.call(this, ...extensions.map(withoutDisabledMarkedExtensions));
  };

  Renderer.prototype.link = safeLink;
  Renderer.prototype.image = safeImage;
}
