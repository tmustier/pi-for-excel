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
import { escapeAttr } from "../../utils/html.js";
import { excelRun } from "../../excel/helpers.js";
import { followCellLink } from "../cell-link.js";
import { copyText } from "../copy-text.js";
import { t } from "../../language/index.js";
import { showToast } from "../toast.js";

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
  return `<div class="pi-md-code"><code-block language="${escapeAttr(language)}" encoding="base64" code="${encoded}"></code-block></div>`;
}

/** Rendered answer text, with real cell separators rather than Markdown pipes. */
export function markdownText(node: Node): string {
  if (node instanceof Text) return node.data;
  if (!(node instanceof Element)) return "";
  if (node.matches("button")) return "";
  if (node instanceof HTMLTableElement) {
    return Array.from(node.rows, (row) => Array.from(row.cells,
      (cell) => cell.innerText.replace(/\s*\n\s*/g, " ").replace(/\t/g, " ").trim(),
    ).join("\t")).join("\n") + "\n\n";
  }
  if (node.localName === "code-block") return node.querySelector("pre")?.textContent ?? "";
  if (node.localName === "br") return "\n";
  const text = Array.from(node.childNodes, markdownText).join("");
  return /^(p|h[1-6]|li|blockquote|pre)$/.test(node.localName) ? text + "\n\n" : text;
}

@customElement("markdown-block")
export class MarkdownBlock extends LitElement {
  @property() content = "";
  @property({ type: Boolean }) isThinking = false;
  @property({ type: Boolean }) escapeHtml = true;
  /** Only completed assistant prose gets workbook links and table actions. */
  @property({ type: Boolean }) answer = false;

  protected override updated(): void {
    if (!this.answer) return;
    for (const button of this.querySelectorAll(".pi-answer-copy")) button.remove();
    for (const table of this.querySelectorAll("table")) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "pi-answer-copy";
      button.textContent = t("messages.copyTable");
      button.addEventListener("click", () => {
        const text = markdownText(table).slice(0, -2);
        void copyText(text).then((outcome) => {
          if (outcome === "copied") showToast(t("code-block.copied"));
        });
      });
      table.parentElement?.append(button);
    }

    const container = this.querySelector(".pi-markdown");
    if (!container || typeof Excel === "undefined") return;
    // One metadata read per completed text block, never per token/reference.
    void excelRun(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      return new Set(sheets.items.map((sheet) => sheet.name.toLowerCase()));
    }).then((sheets) => {
      if (!container.isConnected || container !== this.querySelector(".pi-markdown")) return;
      const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
      const nodes: Text[] = [];
      while (walker.nextNode()) {
        const node = walker.currentNode;
        if (node instanceof Text && !node.parentElement?.closest("a, code, pre, code-block, button")) nodes.push(node);
      }
      for (const node of nodes) {
        const text = node.data;
        const fragment = document.createDocumentFragment();
        let end = 0;
        // Qualified references only: bare Q3/FY24 can be ordinary prose.
        const refs = /(?<![\p{L}\p{N}_!'\[\]])('(?:[^'\n]|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!(\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}))?(?![\p{L}\p{N}_:])/gu;
        for (const match of text.matchAll(refs)) {
          const [address, sheet, start, finish] = match;
          if (!sheet || !start) continue;
          const name = sheet.startsWith("'") ? sheet.slice(1, -1).replace(/''/g, "'") : sheet;
          const valid = [start, finish].every((cell) => {
            if (!cell) return true;
            const letters = cell.replace(/[^A-Za-z]/g, "").toUpperCase();
            const column = Array.from(letters).reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
            return column <= 16384 && Number(cell.replace(/[^0-9]/g, "")) <= 1048576;
          });
          if (!valid || !sheets.has(name.toLowerCase())) continue;
          fragment.append(text.slice(end, match.index));
          const link = document.createElement("a");
          link.className = "pi-cell-ref";
          link.href = "#";
          link.textContent = address;
          link.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            void followCellLink(address);
          });
          fragment.append(link);
          end = match.index + address.length;
        }
        if (end > 0) {
          fragment.append(text.slice(end));
          node.replaceWith(fragment);
        }
      }
    }).catch(() => { /* Workbook metadata unavailable: keep readable plain text. */ });
  }

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    // NOTE: deliberately NOT carrying mini-lit's legacy `markdown-content`
    // class — when pi-web-ui's app.css was still loaded, it styled that class
    // with higher-specificity typography rules that fought our sidebar scale
    // (e.g. `.markdown-content h2` resolving a 40px display token). All
    // typography lives under `markdown-block …` selectors in theme/content
    // CSS instead.
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

    // `breaks`: a single newline is a line break, as people type in chat.
    const parsed = marked.parse(source, { async: false, renderer, breaks: true });

    // Swap <pre><code> blocks for our <code-block> component. The language
    // is whatever marked put in the class (already entity-escaped), so
    // labels such as c++, f# and objective-c keep their code bar.
    let content = parsed.replace(
      /<pre><code class="language-([^"]+)">([\s\S]+?)<\/code><\/pre>/g,
      (_match, language: string, code: string) => codeBlockElement(unescapeHtmlEntities(language), code),
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
