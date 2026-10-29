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

export function iconDOM(lucideIcon: IconNode, size: IconSize = "md", className?: string): SVGElement {
  const px = SIZE_PX[size];
  const element = createElement(lucideIcon, {
    class: `pi-icon pi-icon--${size}${className ? ` ${className}` : ""}`,
    width: px,
    height: px,
  });
  return element;
}

export function icon(lucideIcon: IconNode, size: IconSize = "md", className?: string): TemplateResult {
  return html`${unsafeSVG(iconDOM(lucideIcon, size, className).outerHTML)}`;
}
