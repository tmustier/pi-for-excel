import hljs from "highlight.js/lib/core";
import bash from "highlight.js/lib/languages/bash";
import css from "highlight.js/lib/languages/css";
import javascript from "highlight.js/lib/languages/javascript";
import json from "highlight.js/lib/languages/json";
import python from "highlight.js/lib/languages/python";
import sql from "highlight.js/lib/languages/sql";
import typescript from "highlight.js/lib/languages/typescript";
import xml from "highlight.js/lib/languages/xml";
import { html, LitElement } from "lit";
import { customElement, property, state } from "lit/decorators.js";
import { unsafeHTML } from "lit/directives/unsafe-html.js";
import { Check, Copy } from "lucide";

import { icon } from "../icons.js";
import { t } from "../../language/index.js";
import { decodeBase64Utf8 } from "../../utils/base64-text.js";
import { copyText } from "../copy-text.js";

hljs.registerLanguage("javascript", javascript);
hljs.registerLanguage("typescript", typescript);
hljs.registerLanguage("python", python);
hljs.registerLanguage("html", xml);
hljs.registerLanguage("xml", xml);
hljs.registerLanguage("css", css);
hljs.registerLanguage("json", json);
hljs.registerLanguage("bash", bash);
hljs.registerLanguage("sql", sql);

@customElement("code-block")
export class CodeBlock extends LitElement {
  @property() code = "";
  @property() language = "";
  @property() encoding: "raw" | "base64" = "raw";

  @state() private _copied = false;

  private _copyResetTimer: number | undefined;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  override disconnectedCallback(): void {
    super.disconnectedCallback();
    if (this._copyResetTimer !== undefined) {
      window.clearTimeout(this._copyResetTimer);
      this._copyResetTimer = undefined;
    }
  }

  private _decodedCode(): string {
    if (this.encoding !== "base64") return this.code;
    try {
      return decodeBase64Utf8(this.code);
    } catch {
      return this.code;
    }
  }

  private _copy(): void {
    const code = this._decodedCode();
    void copyText(code)
      .then((outcome) => {
        if (outcome !== "copied") return;
        this._copied = true;
        if (this._copyResetTimer !== undefined) {
          window.clearTimeout(this._copyResetTimer);
        }
        this._copyResetTimer = window.setTimeout(() => {
          this._copied = false;
          this._copyResetTimer = undefined;
        }, 1500);
      })
      .catch(() => {
      });
  }

  override render() {
    const code = this._decodedCode();
    const highlighted =
      this.language && hljs.getLanguage(this.language)
        ? hljs.highlight(code, { language: this.language }).value
        : hljs.highlightAuto(code).value;

    const displayLanguage = this.language || "plaintext";
    const copyLabel = this._copied ? t("code-block.copied") : t("code-block.copy");

    return html`
      <div class="pi-code">
        <div class="pi-code__bar">
          <span class="pi-code__lang">${displayLanguage}</span>
          <button
            type="button"
            class="pi-code__copy"
            title=${t("code-block.copy")}
            @click=${this._copy}
          >
            ${icon(this._copied ? Check : Copy, "xs")}
            <span>${copyLabel}</span>
          </button>
        </div>
        <div class="pi-code__scroll">
          <pre class="pi-code__pre"><code class="hljs language-${displayLanguage}">${unsafeHTML(highlighted)}</code></pre>
        </div>
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    "code-block": CodeBlock;
  }
}
