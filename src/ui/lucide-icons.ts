import { iconDOM } from "./icons.js";
import type { IconNode } from "lucide";
import {
  AlertTriangle,
  Archive,
  Check,
  CircleX,
  ClipboardList,
  Copy,
  FileSpreadsheet,
  FileText,
  FlaskConical,
  Folder,
  FolderOpen,
  Image,
  Info,
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

export function lucide(glyph: IconNode): SVGElement {
  return iconDOM(glyph, "sm");
}

export {
  AlertTriangle,
  Archive,
  Check,
  CircleX,
  ClipboardList,
  Copy,
  FileSpreadsheet,
  FileText,
  FlaskConical,
  Folder,
  FolderOpen,
  Image,
  Info,
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
