/**
 * <markdown-block> — first-party markdown renderer.
 *
 * Replaces @mariozechner/mini-lit's MarkdownBlock (docs/ui-ownership.md),
 * keeping the same tag name, light DOM, and input-hardening behavior:
 * - raw HTML tags in the input are escaped (code spans/fences preserved)
 * - links open in new tabs with rel="noopener noreferrer"
 * - tables get an overflow wrapper
 * - fenced code becomes <code-block> elements (code passed base64-encoded
 *   because it travels through an HTML attribute)
 *
 * Intentionally dropped vs mini-lit: KaTeX math extensions. They were already
 * disabled by src/compat/marked-safety-policy.ts (currency `$` collisions), so
 * this only removes dead weight from the bundle.
 *
 * SECURITY: installMarkedSafetyPatch() (src/compat/marked-safety.ts) patches
 * marked.Renderer.prototype at boot — blocking javascript:/data: links and
 * markdown image network requests. Because we construct `new marked.Renderer()`
 * per render and wrap its (patched) methods, those protections apply here.
 */

import { html, LitElement } from "lit";
import { customElement, property } from "lit/decorators.js";
import { unsafeHTML } from "lit/directives/unsafe-html.js";
import { marked } from "marked";

import { encodeBase64Utf8 } from "../../utils/base64-text.js";

/**
 * Escape raw HTML tags while preserving markdown syntax.
 * Code fences and inline code spans are protected via placeholders.
 */
function escapeRawHtml(content: string): string {
  const codeBlocks: string[] = [];
  let preserved = content.replace(/```[\s\S]*?```|`[^`\n]+`/g, (match) => {
    const index = codeBlocks.length;
    codeBlocks.push(match);
    return `__CODE_BLOCK_${index}__`;
  });

  preserved = preserved
    // Opening tags like <script>, <div>, etc.
    .replace(/<(\w+)([^>]*)>/g, "&lt;$1$2&gt;")
    // Closing tags like </script>, </div>, etc.
    .replace(/<\/(\w+)>/g, "&lt;/$1&gt;")
    // Self-closing tags like <img />, <br/>
    .replace(/<(\w+)([^>]*)\s*\/>/g, "&lt;$1$2/&gt;")
    // Any remaining < that might be part of HTML
    .replace(/<(?![^\s])/g, "&lt;");

  codeBlocks.forEach((block, index) => {
    preserved = preserved.replace(`__CODE_BLOCK_${index}__`, block);
  });

  return preserved;
}

function unescapeHtmlEntities(code: string): string {
  return code
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/g, "'")
    .replace(/&amp;/g, "&"); // Must be last to avoid double-unescaping.
}

function codeBlockElement(language: string, code: string): string {
  // Base64 keeps arbitrary code safe inside an HTML attribute; <code-block>
  // decodes it when encoding="base64" is set.
  const encoded = encodeBase64Utf8(unescapeHtmlEntities(code));
  return `<div class="pi-md-code"><code-block language="${language}" encoding="base64" code="${encoded}"></code-block></div>`;
}

@customElement("markdown-block")
export class MarkdownBlock extends LitElement {
  @property() content = "";
  @property({ type: Boolean }) isThinking = false;
  @property({ type: Boolean }) escapeHtml = true;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    // NOTE: deliberately NOT adding mini-lit's `markdown-content` class —
    // pi-web-ui's app.css styles it with higher-specificity typography rules
    // that fight our sidebar scale (e.g. `.markdown-content h2` resolving
    // var(--text-xl) = 40px from our tokens). All typography lives under
    // `markdown-block …` selectors in theme/content CSS instead.
    this.style.display = "block";
  }

  override render() {
    if (!this.content) {
      return html``;
    }

    const source = this.escapeHtml ? escapeRawHtml(this.content) : this.content;

    // Per-render renderer instance. Its prototype methods were hardened at
    // boot by installMarkedSafetyPatch(), so wrapping them here keeps the
    // link/image protections.
    const renderer = new marked.Renderer();

    const originalLink = renderer.link.bind(renderer);
    renderer.link = (token) =>
      originalLink(token).replace("<a ", '<a target="_blank" rel="noopener noreferrer" ');

    const originalTable = renderer.table.bind(renderer);
    renderer.table = (token) =>
      `<div class="pi-md-table-wrap">${originalTable(token)}</div>`;

    const parsed = marked.parse(source, { async: false, renderer });

    // Swap <pre><code> blocks for our <code-block> component.
    let content = parsed.replace(
      /<pre><code class="language-(\w+)">([\s\S]+?)<\/code><\/pre>/g,
      (_match, language: string, code: string) => codeBlockElement(language, code),
    );
    content = content.replace(
      /<pre><code>([\s\S]+?)<\/code><\/pre>/g,
      (_match, code: string) => codeBlockElement("text", code),
    );

    const containerClass = this.isThinking ? "pi-markdown pi-markdown--thinking" : "pi-markdown";
    return html`<div class=${containerClass}>${unsafeHTML(content)}</div>`;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    "markdown-block": MarkdownBlock;
  }
}
