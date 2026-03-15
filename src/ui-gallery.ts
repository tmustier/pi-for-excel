/**
 * UI Gallery — standalone page for visually verifying components.
 *
 * No Office.js dependency. Loads the same CSS as the real taskpane and
 * renders mock components for agent-browser screenshot verification.
 *
 * Access at: http://localhost:3000/src/ui-gallery.html
 *
 * Each section has a data-gallery attribute for targeted screenshots:
 *   agent-browser screenshot --selector '[data-gallery="tool-cards"]'
 */

// Boot with the same CSS + patches as the real taskpane.
// This imports pi-web-ui/app.css, theme.css, and installs Lit/marked/theme patches.
import "./boot.js";

// Register web components we render
import "./ui/register-components.js";

const galleryRoot = document.getElementById("gallery-root");
if (!galleryRoot) throw new Error("Missing #gallery-root");
const root: HTMLElement = galleryRoot;
root.style.cssText = `
  max-width: 380px;
  margin: 0 auto;
  padding: 16px;
  font-family: var(--font-sans);
  background: var(--background);
  color: var(--foreground);
  min-height: 100vh;
`;

function section(id: string, title: string): HTMLDivElement {
  const el = document.createElement("div");
  el.setAttribute("data-gallery", id);
  el.style.cssText = "margin-bottom: 32px;";

  const heading = document.createElement("h3");
  heading.textContent = title;
  heading.style.cssText = `
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--muted-foreground);
    margin: 0 0 12px;
    padding-bottom: 6px;
    border-bottom: 1px solid var(--alpha-8);
  `;

  el.appendChild(heading);
  root.appendChild(el);
  return el;
}

/* ── 1. Overlay Badges ───────────────────────────────── */

const badgeSection = section("badges", "Overlay Badges");
const badgeRow = document.createElement("div");
badgeRow.style.cssText = "display: flex; gap: 8px; flex-wrap: wrap;";

for (const tone of ["muted", "ok", "warn", "info"]) {
  const badge = document.createElement("span");
  badge.className = `pi-overlay-badge pi-overlay-badge--${tone}`;
  badge.textContent = tone === "muted" ? "Read only" : tone === "ok" ? "Connected" : tone === "warn" ? "Warning" : "Info";
  badgeRow.appendChild(badge);
}
badgeSection.appendChild(badgeRow);

/* ── 2. File Item Rows (with badge) ──────────────────── */

const fileItemSection = section("file-items", "File List Items");

function createMockFileItem(name: string, meta: string, badgeLabel?: string): HTMLButtonElement {
  const row = document.createElement("button");
  row.type = "button";
  row.className = "pi-files-item pi-files-item--muted";
  row.style.cssText = "width: 100%;";

  const icon = document.createElement("span");
  icon.className = "pi-files-item__icon";
  icon.textContent = "📄";

  const info = document.createElement("div");
  info.className = "pi-files-item__info";

  const nameRow = document.createElement("div");
  nameRow.className = "pi-files-item__name-row";

  const nameEl = document.createElement("span");
  nameEl.className = "pi-files-item__name";
  nameEl.textContent = name;
  nameRow.appendChild(nameEl);

  if (badgeLabel) {
    const badge = document.createElement("span");
    badge.className = "pi-overlay-badge pi-overlay-badge--muted";
    badge.textContent = badgeLabel;
    nameRow.appendChild(badge);
  }

  const metaEl = document.createElement("span");
  metaEl.className = "pi-files-item__meta";
  metaEl.textContent = meta;

  info.append(nameRow, metaEl);

  const arrow = document.createElement("span");
  arrow.className = "pi-files-item__arrow";
  arrow.textContent = "›";

  row.append(icon, info, arrow);
  return row;
}

fileItemSection.appendChild(createMockFileItem("cache-observability-baselines.md", "Pi documentation · 2.98 KB", "Read only"));
fileItemSection.appendChild(createMockFileItem("context-management-policy.md", "Pi documentation · 12.0 KB", "Read only"));
fileItemSection.appendChild(createMockFileItem("quarterly-report.xlsx", "1.2 MB · Uploaded · 2h ago"));

/* ── 3. Tool Cards ───────────────────────────────────── */

const toolCardSection = section("tool-cards", "Tool Cards");

function createMockToolCard(state: string, action: string, detail: string): HTMLDivElement {
  const card = document.createElement("div");
  card.className = "pi-tool-card";
  card.setAttribute("data-state", state);
  card.setAttribute("data-tool-name", "fill_formula");

  const header = document.createElement("div");
  header.className = "pi-tool-card__header";

  const toggle = document.createElement("div");
  toggle.className = "pi-tool-card__toggle pi-tool-card__toggle--static";

  const main = document.createElement("span");
  main.className = "pi-tool-card__toggle-main";

  const title = document.createElement("span");
  title.className = "pi-tool-card__title";
  title.innerHTML = `<strong>${action}</strong> <span class="pi-tool-card__detail-text">${detail}</span>`;

  main.appendChild(title);
  toggle.appendChild(main);
  header.appendChild(toggle);
  card.appendChild(header);

  return card;
}

toolCardSection.appendChild(createMockToolCard("complete", "Filled", "'Cash Flow'!D10:L10 — 9 changes"));
toolCardSection.appendChild(createMockToolCard("complete", "Filled", "'Cash Flow'!D13:L13 — 9 changes"));
toolCardSection.appendChild(createMockToolCard("complete", "Filled", "'Cash Flow'!D14:L14 — 9 changes"));
toolCardSection.appendChild(createMockToolCard("error", "Fill", "'Cash Flow'!D15:L15 — error"));

/* ── 4. Tool Card Group ──────────────────────────────── */

const groupSection = section("tool-groups", "Grouped Tool Cards");

const group = document.createElement("div");
group.className = "pi-tool-group";

for (let i = 10; i <= 14; i++) {
  const wrapper = document.createElement("div");
  // Simulate tool-message wrapping
  const card = createMockToolCard("complete", "Filled", `'Cash Flow'!D${i}:L${i} — 9 changes`);
  wrapper.appendChild(card);
  group.appendChild(wrapper);
}
groupSection.appendChild(group);

/* ── 5. Changes Diff Table ───────────────────────────── */

const diffSection = section("diff-table", "Cell Changes Diff Table");

const diffWrap = document.createElement("div");
diffWrap.className = "pi-tool-card__section";
diffWrap.innerHTML = `
  <div class="pi-tool-card__section-label">Changes (9)</div>
  <div class="pi-tool-card__diff">
    <table class="pi-tool-card__diff-table">
      <thead>
        <tr><th>Cell</th><th>Before</th><th>After</th></tr>
      </thead>
      <tbody>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">D10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$125,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =C19+C24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$130,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =D19+D24</div>
          </td>
        </tr>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">E10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$130,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =D19+D24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$135,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =E19+E24</div>
          </td>
        </tr>
        <tr>
          <td class="pi-tool-card__diff-cell"><span class="pi-cell-ref">F10</span></td>
          <td>
            <div class="pi-tool-card__diff-value">$135,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =E19+E24</div>
          </td>
          <td>
            <div class="pi-tool-card__diff-value">$140,000</div>
            <div class="pi-tool-card__diff-formula">ƒ =F19+F24</div>
          </td>
        </tr>
      </tbody>
    </table>
  </div>
`;
diffSection.appendChild(diffWrap);

/* ── 6. Text Preview (file detail) ───────────────────── */

const previewSection = section("text-preview", "File Text Preview");

const preview = document.createElement("div");
preview.className = "pi-files-detail-preview pi-files-detail-preview--text";

const sampleLines = [
  "# Context Management Policy",
  "",
  "**Status:** Active policy (2026-02-12)",
  "**Scope:** How Pi for Excel builds and manages context",
  "",
  "---",
  "",
  "## Why this exists",
  "",
  "We optimize for **answer quality and reliability** across multi-turn sessions.",
  "",
  "In practice, quality drops when we blindly stuff context or let it grow unbounded.",
  "",
  "## Core principles",
  "",
  "1. **Minimal viable context** — include only what improves this turn.",
  "2. **Freshness over volume** — recent state > historical state.",
  "3. **Structured disclosure** — progressive detail, not a wall of text.",
  "4. **Cache-friendly ordering** — static prefix, dynamic tail.",
  "5. **Bounded growth** — auto-compact before hitting limits.",
];

sampleLines.forEach((line, i) => {
  const lineRow = document.createElement("div");
  lineRow.className = "pi-files-detail-preview__line";
  if (i === 0) lineRow.style.paddingTop = "8px";
  if (i === sampleLines.length - 1) lineRow.style.paddingBottom = "8px";

  const ln = document.createElement("span");
  ln.className = "pi-files-detail-preview__ln";
  ln.textContent = String(i + 1);

  const code = document.createElement("span");
  code.className = "pi-files-detail-preview__code";
  code.textContent = line;

  lineRow.append(ln, code);
  preview.appendChild(lineRow);
});
previewSection.appendChild(preview);

/* ── 7. Action Buttons ───────────────────────────────── */

const buttonsSection = section("buttons", "Overlay Buttons");

const btnRow = document.createElement("div");
btnRow.className = "pi-files-detail-actions";

for (const [label, cls] of [
  ["Open ↗", "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact"],
  ["Download", "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact"],
  ["Delete", "pi-overlay-btn pi-overlay-btn--danger pi-overlay-btn--compact"],
] as const) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = cls;
  btn.textContent = label;
  btnRow.appendChild(btn);
}
buttonsSection.appendChild(btnRow);

/* ── 8. Toast variants ───────────────────────────────── */

const toastSection = section("toasts", "Toast Notifications");

for (const [msg, classes] of [
  ["Closed Chat 3", "pi-toast visible pi-toast--action"],
  ["Tab name reset", "pi-toast visible"],
  ["Could not save", "pi-toast visible pi-toast--error"],
] as const) {
  const toast = document.createElement("div");
  toast.className = classes;
  toast.style.cssText = "position: relative; top: 0; left: 0; transform: none; opacity: 1; pointer-events: auto; margin-bottom: 8px;";

  const content = document.createElement("div");
  content.className = "pi-toast__content";

  const message = document.createElement("span");
  message.className = "pi-toast__message";
  message.textContent = msg;

  content.appendChild(message);

  if (classes.includes("pi-toast--action")) {
    const action = document.createElement("button");
    action.type = "button";
    action.className = "pi-toast__action";
    action.textContent = "Undo";
    content.appendChild(action);
  }

  toast.appendChild(content);
  toastSection.appendChild(toast);
}

/* ── 9. Markdown rendering (font test) ───────────────── */

const mdSection = section("markdown", "Markdown Rendering (font consistency)");

const mdBlock = document.createElement("markdown-block") as HTMLElement & { content: string };
mdBlock.content = `The formula is \`=IF(C$4-Assumptions!$B$10+1=Assumptions!$B$49,...)\` — C4 = calendaryear (2025 for Year 1, 2031 for Year 7).

Assumptions!B10 = 2025 (start year) − Assumptions!B$49 = 7

So for Year 7 (column I, calendar year 2031): 2031 – 2025 + 1 = 7 ✓`;

mdSection.appendChild(mdBlock);

console.log("[ui-gallery] Rendered all sections");
