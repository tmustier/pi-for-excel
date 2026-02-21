/**
 * Thin wrapper around Lucide + mini-lit iconDOM for imperative DOM code.
 *
 * Overlay builders that construct DOM via `document.createElement` cannot
 * use Lit's `html` tagged template. This module re-exports `iconDOM` and
 * the Lucide glyphs used across overlay dialogs so each file doesn't need
 * to duplicate imports.
 */

import { iconDOM } from "@mariozechner/mini-lit";
import type { IconNode } from "lucide";
import {
  AlertTriangle,
  Check,
  ClipboardList,
  Copy,
  FileSpreadsheet,
  FileText,
  Folder,
  FolderOpen,
  Image,
  Link,
  NotebookPen,
  Package,
  Paperclip,
  Puzzle,
  Search,
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
  Check,
  ClipboardList,
  Copy,
  FileSpreadsheet,
  FileText,
  Folder,
  FolderOpen,
  Image,
  Link,
  NotebookPen,
  Package,
  Paperclip,
  Puzzle,
  Search,
  Terminal,
  Upload,
  Zap,
};
