/**
 * Thin wrapper around Lucide + mini-lit iconDOM for imperative DOM code.
 *
 * Overlay builders that construct DOM via `document.createElement` cannot
 * use Lit's `html` tagged template. This module re-exports `iconDOM` and
 * the Lucide glyphs used across overlay dialogs so each file doesn't need
 * to duplicate imports.
 */

import { iconDOM } from "./icons.js";
import type { IconNode } from "lucide";
import {
  AlertTriangle,
  Archive,
  Check,
  ClipboardList,
  Copy,
  FileSpreadsheet,
  FileText,
  FlaskConical,
  Folder,
  FolderOpen,
  Image,
  Keyboard,
  Link,
  NotebookPen,
  Package,
  Paperclip,
  Plug,
  Puzzle,
  Ruler,
  Search,
  Server,
  ShieldCheck,
  Terminal,
  Upload,
  Zap,
} from "lucide";

export type { IconNode };

/** Create a 16×16 SVG element for use in imperative DOM code. */
export function lucide(glyph: IconNode): SVGElement {
  return iconDOM(glyph, "sm");
}

export {
  AlertTriangle,
  Archive,
  Check,
  ClipboardList,
  Copy,
  FileSpreadsheet,
  FileText,
  FlaskConical,
  Folder,
  FolderOpen,
  Image,
  Keyboard,
  Link,
  NotebookPen,
  Package,
  Paperclip,
  Plug,
  Puzzle,
  Ruler,
  Search,
  Server,
  ShieldCheck,
  Terminal,
  Upload,
  Zap,
};
