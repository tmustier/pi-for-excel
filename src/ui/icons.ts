/**
 * First-party Lucide icon helpers.
 *
 * Replaces @mariozechner/mini-lit's `icon`/`iconDOM` (which sized icons via
 * Tailwind utility classes). We size via explicit SVG width/height attributes
 * so icons render correctly without any utility CSS. See docs/ui-ownership.md.
 */

import { html, type TemplateResult } from "lit";
import { unsafeSVG } from "lit/directives/unsafe-svg.js";
import { createElement, type IconNode } from "lucide";

export type IconSize = "xs" | "sm" | "md" | "lg" | "xl";

const SIZE_PX: Record<IconSize, number> = {
  xs: 12,
  sm: 16,
  md: 20,
  lg: 24,
  xl: 32,
};

/** Create a sized Lucide SVG element for imperative DOM code. */
export function iconDOM(lucideIcon: IconNode, size: IconSize = "md", className?: string): SVGElement {
  const px = SIZE_PX[size];
  const element = createElement(lucideIcon, {
    class: `pi-icon pi-icon--${size}${className ? ` ${className}` : ""}`,
    width: px,
    height: px,
  });
  return element;
}

/** Create a sized Lucide icon template for Lit renders. */
export function icon(lucideIcon: IconNode, size: IconSize = "md", className?: string): TemplateResult {
  return html`${unsafeSVG(iconDOM(lucideIcon, size, className).outerHTML)}`;
}
