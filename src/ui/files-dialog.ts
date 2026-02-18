/**
 * Files workspace dialog.
 */

import { base64ToBytes } from "../files/encoding.js";
import { formatBytes } from "../files/mime.js";
import {
  FILES_WORKSPACE_CHANGED_EVENT,
  type WorkspaceBackendStatus,
  type WorkspaceFileEntry,
  type WorkspaceFileLocationKind,
  type WorkspaceFileWorkbookTag,
} from "../files/types.js";
import { type FilesWorkspaceAuditContext, getFilesWorkspace } from "../files/workspace.js";
import { getErrorMessage } from "../utils/errors.js";
import { requestConfirmationDialog } from "./confirm-dialog.js";
import {
  buildFilesDialogSections,
  type FilesDialogFolderGroup,
  isAgentWrittenNotesFilePath,
  isFilesDialogBuiltInDoc,
  normalizeFilesDialogFilterText,
  resolveFilesDialogBadge,
  resolveFilesDialogConnectFolderButtonState,
  resolveFilesDialogSourceLabel,
  sectionTotalCount,
} from "./files-dialog-filtering.js";
import { buildFilesDialogStatusMessage } from "./files-dialog-status.js";
import {
  closeOverlayById,
  createOverlayCloseButton,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { FILES_WORKSPACE_OVERLAY_ID } from "./overlay-ids.js";
import { resolveSafeFilesDialogBlobMimeType } from "./files-dialog-mime.js";
import { resolveRenameDestinationPath } from "./files-dialog-paths.js";
import { requestTextInputDialog } from "./text-input-dialog.js";
import { showToast } from "./toast.js";
import type { IconContent } from "./extensions-hub-components.js";
import {
  lucide,
  AlertTriangle,
  ClipboardList,
  FileSpreadsheet,
  FileText,
  Folder,
  FolderOpen,
  Image,
  Link,
  NotebookPen,
  Paperclip,
  Search,
  Upload,
} from "./lucide-icons.js";

const OVERLAY_ID = FILES_WORKSPACE_OVERLAY_ID;
const TEXT_PREVIEW_MAX_LINES = 50;

const DIALOG_AUDIT_CONTEXT: FilesWorkspaceAuditContext = {
  actor: "user",
  source: "files-dialog",
};

interface DetailPreviewResult {
  element: HTMLElement;
  previewTruncated: boolean;
  objectUrl: string | null;
}

interface FilesDialogFileRef {
  path: string;
  locationKind: WorkspaceFileLocationKind;
}

function resolveFileLocationKind(file: WorkspaceFileEntry): WorkspaceFileLocationKind {
  if (file.locationKind) {
    return file.locationKind;
  }

  if (isFilesDialogBuiltInDoc(file)) {
    return "builtin-doc";
  }

  return "workspace";
}

function toFileRef(file: WorkspaceFileEntry): FilesDialogFileRef {
  return {
    path: file.path,
    locationKind: resolveFileLocationKind(file),
  };
}

function fileMatchesRef(file: WorkspaceFileEntry, ref: FilesDialogFileRef): boolean {
  return file.path === ref.path && resolveFileLocationKind(file) === ref.locationKind;
}

function formatRelativeDate(timestamp: number): string {
  const now = Date.now();
  const diff = now - timestamp;

  if (diff < 60_000) return "just now";
  if (diff < 3_600_000) return `${Math.round(diff / 60_000)}m ago`;
  if (diff < 86_400_000) return `${Math.round(diff / 3_600_000)}h ago`;
  if (diff < 604_800_000) return `${Math.round(diff / 86_400_000)}d ago`;
  return new Date(timestamp).toLocaleDateString();
}

function isImageMimeType(mimeType: string): boolean {
  return mimeType.toLowerCase().startsWith("image/");
}

function isPdfMimeType(mimeType: string): boolean {
  return mimeType.trim().toLowerCase() === "application/pdf";
}

function hasOneOfExtensions(path: string, extensions: readonly string[]): boolean {
  const lowerPath = path.toLowerCase();
  return extensions.some((extension) => lowerPath.endsWith(extension));
}

function resolveFileIcon(file: WorkspaceFileEntry): SVGElement {
  if (isFilesDialogBuiltInDoc(file)) {
    return lucide(ClipboardList);
  }

  if (isImageMimeType(file.mimeType)) {
    return lucide(Image);
  }

  if (hasOneOfExtensions(file.path, [".csv", ".xlsx", ".xls"])) {
    return lucide(FileSpreadsheet);
  }

  if (isAgentWrittenNotesFilePath(file.path)) {
    return lucide(NotebookPen);
  }

  return lucide(FileText);
}

function buildFileMetaLine(file: WorkspaceFileEntry): string {
  const sourceLabel = resolveFilesDialogSourceLabel(file);
  if (isFilesDialogBuiltInDoc(file)) {
    return `${sourceLabel} · ${formatBytes(file.size)}`;
  }

  return `${formatBytes(file.size)} · ${sourceLabel} · ${formatRelativeDate(file.modifiedAt)}`;
}

function resolveDetailTypeName(file: WorkspaceFileEntry): string {
  if (isImageMimeType(file.mimeType)) {
    return "Image";
  }

  if (isPdfMimeType(file.mimeType)) {
    return "PDF";
  }

  if (hasOneOfExtensions(file.path, [".csv"])) {
    return "CSV";
  }

  if (file.kind === "text") {
    return "Text";
  }

  return "Binary";
}

function buildDetailSubtitle(args: {
  file: WorkspaceFileEntry;
  previewTruncated: boolean;
}): string {
  const sourceLabel = resolveFilesDialogSourceLabel(args.file);
  const pieces = [
    resolveDetailTypeName(args.file),
    formatBytes(args.file.size),
    sourceLabel,
  ];

  if (!isFilesDialogBuiltInDoc(args.file)) {
    pieces.push(formatRelativeDate(args.file.modifiedAt));
  }

  const base = pieces.join(" · ");
  if (args.previewTruncated) {
    return `${base} (preview truncated)`;
  }

  return base;
}

function createChevronIcon(): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.classList.add("pi-files-section-head__chevron");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.5");

  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", "M4 6l4 4 4-4");
  svg.appendChild(path);

  return svg;
}

function createInfoCallout(args: {
  icon: IconContent;
  body: Array<HTMLElement | string>;
}): HTMLDivElement {
  const callout = document.createElement("div");
  callout.className = "pi-callout pi-callout--info";

  const iconEl = document.createElement("span");
  iconEl.className = "pi-callout__icon";
  if (typeof args.icon === "string") {
    iconEl.textContent = args.icon;
  } else {
    iconEl.appendChild(args.icon);
  }

  const body = document.createElement("div");
  body.className = "pi-callout__body";

  for (const item of args.body) {
    if (typeof item === "string") {
      body.append(item);
      continue;
    }

    body.appendChild(item);
  }

  callout.append(iconEl, body);
  return callout;
}

function createBinaryPreview(args: {
  file: WorkspaceFileEntry;
  icon: IconContent;
  label: string;
}): HTMLDivElement {
  const preview = document.createElement("div");
  preview.className = "pi-files-detail-preview pi-files-detail-preview--binary";

  const placeholder = document.createElement("div");
  placeholder.className = "pi-files-detail-preview__placeholder";

  const iconEl = document.createElement("span");
  iconEl.className = "pi-files-detail-preview__placeholder-icon";
  if (typeof args.icon === "string") {
    iconEl.textContent = args.icon;
  } else {
    iconEl.appendChild(args.icon);
  }

  const label = document.createElement("span");
  label.className = "pi-files-detail-preview__placeholder-label";
  label.textContent = args.label;

  const size = document.createElement("span");
  size.className = "pi-files-detail-preview__placeholder-size";
  size.textContent = formatBytes(args.file.size);

  placeholder.append(iconEl, label, size);
  preview.appendChild(placeholder);
  return preview;
}

function createTextPreview(text: string, truncated: boolean): {
  element: HTMLDivElement;
  hasMoreLines: boolean;
} {
  const preview = document.createElement("div");
  preview.className = "pi-files-detail-preview pi-files-detail-preview--text";

  const lines = text.replaceAll("\r\n", "\n").split("\n");
  const visibleLines = lines.slice(0, TEXT_PREVIEW_MAX_LINES);

  visibleLines.forEach((line, index) => {
    const lineRow = document.createElement("div");
    lineRow.className = "pi-files-detail-preview__line";

    const lineNumber = document.createElement("span");
    lineNumber.className = "pi-files-detail-preview__ln";
    lineNumber.textContent = String(index + 1);

    const code = document.createElement("span");
    code.className = "pi-files-detail-preview__code";
    code.textContent = line;

    lineRow.append(lineNumber, code);
    preview.appendChild(lineRow);
  });

  const hasMoreLines = lines.length > TEXT_PREVIEW_MAX_LINES;
  if (hasMoreLines && !truncated) {
    const fadeLine = document.createElement("div");
    fadeLine.className = "pi-files-detail-preview__line pi-files-detail-preview__line--fade";

    const lineNumber = document.createElement("span");
    lineNumber.className = "pi-files-detail-preview__ln";
    lineNumber.textContent = String(TEXT_PREVIEW_MAX_LINES + 1);

    const code = document.createElement("span");
    code.className = "pi-files-detail-preview__code";

    fadeLine.append(lineNumber, code);
    preview.appendChild(fadeLine);
  }

  return {
    element: preview,
    hasMoreLines,
  };
}

function createUploadActionButton(args: {
  icon: IconContent;
  label: string;
}): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact";

  const iconEl = document.createElement("span");
  iconEl.className = "pi-files-actions__icon";
  if (typeof args.icon === "string") {
    iconEl.textContent = args.icon;
  } else {
    iconEl.appendChild(args.icon);
  }

  button.append(iconEl, ` ${args.label}`);
  return button;
}

function createEmptyState(onUpload: () => void): HTMLDivElement {
  const empty = document.createElement("div");
  empty.className = "pi-files-empty";

  const emptyIcon = document.createElement("div");
  emptyIcon.className = "pi-files-empty__icon";
  emptyIcon.appendChild(lucide(FileText));

  const title = document.createElement("div");
  title.className = "pi-files-empty__title";
  title.textContent = "Give Pi more context";

  const description = document.createElement("p");
  description.className = "pi-files-empty__desc";
  description.textContent = "Upload documents, data, or reference material to help Pi give better answers.";

  const uploadButton = document.createElement("button");
  uploadButton.type = "button";
  uploadButton.className = "pi-overlay-btn pi-overlay-btn--primary pi-overlay-btn--compact";
  uploadButton.textContent = "Upload files";
  uploadButton.addEventListener("click", onUpload);

  const hint = document.createElement("p");
  hint.className = "pi-files-empty__hint";
  hint.textContent = "Files are stored locally in your browser.";

  empty.append(emptyIcon, title, description, uploadButton, hint);
  return empty;
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
}

function closeWindowSafely(windowHandle: Window | null): void {
  if (!windowHandle || windowHandle.closed) {
    return;
  }

  try {
    windowHandle.close();
  } catch {
    // Ignore close errors.
  }
}

function openBlobInNewTab(blob: Blob, pendingWindow: Window | null): void {
  const url = URL.createObjectURL(blob);
  let opened = false;

  if (pendingWindow && !pendingWindow.closed) {
    try {
      pendingWindow.location.replace(url);
      opened = true;
    } catch {
      closeWindowSafely(pendingWindow);
    }
  }

  if (!opened) {
    opened = window.open(url, "_blank") !== null;
  }

  if (!opened) {
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.target = "_blank";
    anchor.rel = "noopener noreferrer";
    anchor.style.display = "none";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }

  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 60_000);
}

function createWorkbookTagCallout(workbookTag: WorkspaceFileWorkbookTag): HTMLDivElement {
  const strong = document.createElement("strong");
  strong.textContent = workbookTag.workbookLabel;

  return createInfoCallout({
    icon: lucide(Link),
    body: ["Tagged to ", strong, " — included when that workbook is open."],
  });
}

export async function showFilesWorkspaceDialog(): Promise<void> {
  if (closeOverlayById(OVERLAY_ID)) {
    return;
  }

  const workspace = getFilesWorkspace();

  const dialog = createOverlayDialog({
    overlayId: OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--m pi-files-dialog",
  });

  const closeOverlay = dialog.close;

  const listHeaderElements = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close files",
    title: "Files",
    subtitle: "Documents available to Pi",
  });

  const detailHeader = document.createElement("div");
  detailHeader.className = "pi-overlay-header";
  detailHeader.hidden = true;

  const detailTitleContainer = document.createElement("div");
  detailTitleContainer.className = "pi-files-detail-title";

  const detailBackButton = document.createElement("button");
  detailBackButton.type = "button";
  detailBackButton.className = "pi-files-detail__back";
  detailBackButton.setAttribute("aria-label", "Back to file list");
  detailBackButton.textContent = "←";

  const detailTitleWrap = document.createElement("div");
  detailTitleWrap.className = "pi-overlay-title-wrap";

  const detailTitle = document.createElement("h2");
  detailTitle.className = "pi-overlay-title pi-overlay-title--sm";

  const detailSubtitle = document.createElement("p");
  detailSubtitle.className = "pi-overlay-subtitle";

  detailTitleWrap.append(detailTitle, detailSubtitle);
  detailTitleContainer.append(detailBackButton, detailTitleWrap);

  const detailCloseButton = createOverlayCloseButton({
    onClose: closeOverlay,
    label: "Close files",
  });

  detailHeader.append(detailTitleContainer, detailCloseButton);

  const actionsRow = document.createElement("div");
  actionsRow.className = "pi-files-actions";

  const uploadButton = createUploadActionButton({
    icon: lucide(Upload),
    label: "Upload",
  });

  const connectFolderButton = createUploadActionButton({
    icon: lucide(FolderOpen),
    label: "Connect folder",
  });
  connectFolderButton.hidden = true;
  connectFolderButton.disabled = true;

  actionsRow.append(uploadButton, connectFolderButton);

  const hiddenInput = document.createElement("input");
  hiddenInput.type = "file";
  hiddenInput.multiple = true;
  hiddenInput.hidden = true;

  const listBody = document.createElement("div");
  listBody.className = "pi-overlay-body";

  const filterHost = document.createElement("div");

  const filterWrap = document.createElement("div");
  filterWrap.className = "pi-files-filter";

  const filterIcon = document.createElement("span");
  filterIcon.className = "pi-files-filter__icon";
  filterIcon.appendChild(lucide(Search));

  const filterInput = document.createElement("input");
  filterInput.type = "text";
  filterInput.className = "pi-files-filter__input";
  filterInput.placeholder = "Filter files…";
  filterInput.addEventListener("input", () => {
    filterText = filterInput.value;
    renderListView();
  });

  filterWrap.append(filterIcon, filterInput);
  filterHost.appendChild(filterWrap);

  const sectionsHost = document.createElement("div");
  listBody.append(filterHost, sectionsHost);

  const footer = document.createElement("div");
  footer.className = "pi-files-footer";

  const detailBody = document.createElement("div");
  detailBody.className = "pi-overlay-body";
  detailBody.hidden = true;

  dialog.card.append(
    listHeaderElements.header,
    detailHeader,
    actionsRow,
    hiddenInput,
    listBody,
    footer,
    detailBody,
  );

  let backendStatus: WorkspaceBackendStatus | null = null;
  let allFiles: WorkspaceFileEntry[] = [];
  let currentView: "list" | "detail" = "list";
  let detailFileRef: FilesDialogFileRef | null = null;
  let filterText = "";
  let activePreviewObjectUrl: string | null = null;
  let detailRenderVersion = 0;
  const collapsedSections = new Set<string>();

  const revokePreviewObjectUrl = (): void => {
    if (!activePreviewObjectUrl) {
      return;
    }

    URL.revokeObjectURL(activePreviewObjectUrl);
    activePreviewObjectUrl = null;
  };

  const findFileByRef = (ref: FilesDialogFileRef): WorkspaceFileEntry | null => {
    return allFiles.find((file) => fileMatchesRef(file, ref)) ?? null;
  };

  const refreshWorkspaceState = async (): Promise<void> => {
    const [backend, files] = await Promise.all([
      workspace.getBackendStatus(),
      workspace.listFiles(),
    ]);

    backendStatus = backend;
    allFiles = files;
  };

  const setView = (view: "list" | "detail"): void => {
    currentView = view;
    const showList = view === "list";

    listHeaderElements.header.hidden = !showList;
    actionsRow.hidden = !showList;
    listBody.hidden = !showList;
    footer.hidden = !showList;

    detailHeader.hidden = showList;
    detailBody.hidden = showList;
  };

  const showListView = (): void => {
    detailFileRef = null;
    detailRenderVersion += 1;
    revokePreviewObjectUrl();
    setView("list");
  };

  const openFileInBrowser = async (file: WorkspaceFileEntry): Promise<void> => {
    const fileRef = toFileRef(file);
    const pendingWindow = window.open("", "_blank");

    try {
      if (file.kind === "text") {
        const result = await workspace.readFile(file.path, {
          mode: "text",
          maxChars: 16_000_000,
          audit: DIALOG_AUDIT_CONTEXT,
          locationKind: fileRef.locationKind,
        });

        if (result.text === undefined || result.truncated) {
          throw new Error("File is too large to open in a browser tab.");
        }

        const blob = new Blob([result.text], {
          type: resolveSafeFilesDialogBlobMimeType(file.mimeType || "text/plain"),
        });

        openBlobInNewTab(blob, pendingWindow);
        return;
      }

      const result = await workspace.readFile(file.path, {
        mode: "base64",
        maxChars: 16_000_000,
        audit: DIALOG_AUDIT_CONTEXT,
        locationKind: fileRef.locationKind,
      });

      if (!result.base64 || result.truncated) {
        throw new Error("File is too large to open in a browser tab.");
      }

      const bytes = base64ToBytes(result.base64);
      const blob = new Blob([toArrayBuffer(bytes)], {
        type: resolveSafeFilesDialogBlobMimeType(file.mimeType),
      });

      openBlobInNewTab(blob, pendingWindow);
    } catch (error: unknown) {
      closeWindowSafely(pendingWindow);
      throw error;
    }
  };

  const createDetailActions = (file: WorkspaceFileEntry): HTMLDivElement => {
    const fileRef = toFileRef(file);
    const actions = document.createElement("div");
    actions.className = "pi-files-detail-actions";

    const openButton = document.createElement("button");
    openButton.type = "button";
    openButton.className = file.kind === "text"
      ? "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact"
      : "pi-overlay-btn pi-overlay-btn--primary pi-overlay-btn--compact";
    openButton.textContent = "Open ↗";
    openButton.addEventListener("click", () => {
      void openFileInBrowser(file).catch((error: unknown) => {
        showToast(`Open failed: ${getErrorMessage(error)}`);
      });
    });

    const downloadButton = document.createElement("button");
    downloadButton.type = "button";
    downloadButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact";
    downloadButton.textContent = "Download";
    downloadButton.addEventListener("click", () => {
      void workspace.downloadFile(file.path, {
        locationKind: fileRef.locationKind,
      }).catch((error: unknown) => {
        showToast(`Download failed: ${getErrorMessage(error)}`);
      });
    });

    actions.append(openButton, downloadButton);

    const isReadOnly = file.readOnly || isFilesDialogBuiltInDoc(file);
    if (isReadOnly) {
      return actions;
    }

    const renameButton = document.createElement("button");
    renameButton.type = "button";
    renameButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact";
    renameButton.textContent = "Rename";
    renameButton.addEventListener("click", () => {
      void (async () => {
        const nextPathInput = await requestTextInputDialog({
          title: "Rename file",
          message: `${file.path} — leave off the extension to keep it.`,
          initialValue: file.path,
          placeholder: "folder/file.ext",
          confirmLabel: "Rename",
          cancelLabel: "Cancel",
          restoreFocusOnClose: false,
        });

        if (nextPathInput === null) {
          return;
        }

        const nextPath = resolveRenameDestinationPath(file.path, nextPathInput);
        if (nextPath === file.path) {
          return;
        }

        await workspace.renameFile(file.path, nextPath, {
          audit: DIALOG_AUDIT_CONTEXT,
          locationKind: fileRef.locationKind,
        });

        showToast(`Renamed to ${nextPath}.`);

        await refreshWorkspaceState();
        renderListView();
        await showDetailView({
          path: nextPath,
          locationKind: fileRef.locationKind,
        });
      })().catch((error: unknown) => {
        showToast(`Rename failed: ${getErrorMessage(error)}`);
      });
    });

    const spacer = document.createElement("div");
    spacer.className = "pi-files-detail-actions__spacer";

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "pi-overlay-btn pi-overlay-btn--danger pi-overlay-btn--compact";
    deleteButton.textContent = "Delete";
    deleteButton.addEventListener("click", () => {
      void (async () => {
        const confirmed = await requestConfirmationDialog({
          title: "Delete file?",
          message: file.path,
          confirmLabel: "Delete",
          cancelLabel: "Cancel",
          confirmButtonTone: "danger",
          restoreFocusOnClose: false,
        });

        if (!confirmed) {
          return;
        }

        await workspace.deleteFile(file.path, {
          audit: DIALOG_AUDIT_CONTEXT,
          locationKind: fileRef.locationKind,
        });

        showToast(`Deleted ${file.name}.`);

        await refreshWorkspaceState();
        renderListView();
        showListView();
      })().catch((error: unknown) => {
        showToast(`Delete failed: ${getErrorMessage(error)}`);
      });
    });

    actions.append(renameButton, spacer, deleteButton);
    return actions;
  };

  const buildDetailPreview = async (file: WorkspaceFileEntry): Promise<DetailPreviewResult> => {
    const fileRef = toFileRef(file);

    if (file.kind === "text") {
      const result = await workspace.readFile(file.path, {
        mode: "text",
        maxChars: 50_000,
        audit: DIALOG_AUDIT_CONTEXT,
        locationKind: fileRef.locationKind,
      });

      const preview = createTextPreview(result.text ?? "", result.truncated === true);
      return {
        element: preview.element,
        previewTruncated: result.truncated === true,
        objectUrl: null,
      };
    }

    if (isImageMimeType(file.mimeType)) {
      const result = await workspace.readFile(file.path, {
        mode: "base64",
        maxChars: 8_000_000,
        audit: DIALOG_AUDIT_CONTEXT,
        locationKind: fileRef.locationKind,
      });

      if (!result.base64 || result.truncated) {
        return {
          element: createBinaryPreview({
            file,
            icon: lucide(Image),
            label: "Image preview unavailable",
          }),
          previewTruncated: false,
          objectUrl: null,
        };
      }

      const bytes = base64ToBytes(result.base64);
      const url = URL.createObjectURL(new Blob([toArrayBuffer(bytes)], { type: file.mimeType }));

      const preview = document.createElement("div");
      preview.className = "pi-files-detail-preview pi-files-detail-preview--image";

      const image = document.createElement("img");
      image.src = url;
      image.alt = file.name;
      preview.appendChild(image);

      return {
        element: preview,
        previewTruncated: false,
        objectUrl: url,
      };
    }

    if (isPdfMimeType(file.mimeType)) {
      return {
        element: createBinaryPreview({
          file,
          icon: lucide(FileText),
          label: "PDF document",
        }),
        previewTruncated: false,
        objectUrl: null,
      };
    }

    return {
      element: createBinaryPreview({
        file,
        icon: lucide(Paperclip),
        label: "Binary file",
      }),
      previewTruncated: false,
      objectUrl: null,
    };
  };

  const renderDetailView = async (fileRef: FilesDialogFileRef): Promise<void> => {
    const file = findFileByRef(fileRef);
    if (!file) {
      showListView();
      renderListView();
      return;
    }

    const renderVersion = ++detailRenderVersion;
    revokePreviewObjectUrl();

    detailTitle.textContent = file.name;
    detailSubtitle.textContent = buildDetailSubtitle({
      file,
      previewTruncated: false,
    });

    const nodes: HTMLElement[] = [];

    if (file.workbookTag) {
      nodes.push(createWorkbookTagCallout(file.workbookTag));
    }

    if (isFilesDialogBuiltInDoc(file)) {
      nodes.push(createInfoCallout({
        icon: lucide(ClipboardList),
        body: ["Built-in documentation — read only. Pi references this automatically."],
      }));
    }

    let previewResult: DetailPreviewResult;
    try {
      previewResult = await buildDetailPreview(file);
    } catch (error: unknown) {
      previewResult = {
        element: createBinaryPreview({
          file,
          icon: lucide(AlertTriangle),
          label: `Preview unavailable: ${getErrorMessage(error)}`,
        }),
        previewTruncated: false,
        objectUrl: null,
      };
    }

    if (
      renderVersion !== detailRenderVersion ||
      currentView !== "detail" ||
      !detailFileRef ||
      detailFileRef.path !== fileRef.path ||
      detailFileRef.locationKind !== fileRef.locationKind
    ) {
      if (previewResult.objectUrl) {
        URL.revokeObjectURL(previewResult.objectUrl);
      }

      return;
    }

    if (previewResult.objectUrl) {
      activePreviewObjectUrl = previewResult.objectUrl;
    }

    detailSubtitle.textContent = buildDetailSubtitle({
      file,
      previewTruncated: previewResult.previewTruncated,
    });

    nodes.push(previewResult.element, createDetailActions(file));
    detailBody.replaceChildren(...nodes);
  };

  const showDetailView = async (fileRef: FilesDialogFileRef): Promise<void> => {
    detailFileRef = fileRef;
    setView("detail");
    await renderDetailView(fileRef);
  };

  const createFileItem = (file: WorkspaceFileEntry): HTMLButtonElement => {
    const row = document.createElement("button");
    row.type = "button";
    row.className = `pi-files-item${isFilesDialogBuiltInDoc(file) ? " pi-files-item--muted" : ""}`;

    const icon = document.createElement("span");
    icon.className = "pi-files-item__icon";
    icon.appendChild(resolveFileIcon(file));

    const info = document.createElement("div");
    info.className = "pi-files-item__info";

    const nameRow = document.createElement("div");
    nameRow.className = "pi-files-item__name-row";

    const name = document.createElement("span");
    name.className = "pi-files-item__name";
    name.textContent = file.name;
    name.title = file.path;
    nameRow.appendChild(name);

    const badge = resolveFilesDialogBadge(file);
    if (badge) {
      const badgeElement = document.createElement("span");
      badgeElement.className = `pi-overlay-badge pi-overlay-badge--${badge.tone}`;
      badgeElement.textContent = badge.label;
      nameRow.appendChild(badgeElement);
    }

    const meta = document.createElement("span");
    meta.className = "pi-files-item__meta";
    meta.textContent = buildFileMetaLine(file);

    info.append(nameRow, meta);

    const arrow = document.createElement("span");
    arrow.className = "pi-files-item__arrow";
    arrow.textContent = "›";

    row.append(icon, info, arrow);
    row.addEventListener("click", () => {
      void showDetailView(toFileRef(file));
    });

    return row;
  };

  const createFolderGroup = (folder: FilesDialogFolderGroup, collapseKey: string): HTMLDivElement => {
    const group = document.createElement("div");
    group.className = "pi-files-folder-group";

    const header = document.createElement("button");
    header.type = "button";
    header.className = "pi-files-folder";

    const iconEl = document.createElement("span");
    iconEl.className = "pi-files-folder__icon";
    iconEl.appendChild(lucide(Folder));

    const nameEl = document.createElement("span");
    nameEl.className = "pi-files-folder__name";
    nameEl.textContent = folder.name;

    const countEl = document.createElement("span");
    countEl.className = "pi-files-folder__count";
    const n = folder.files.length;
    countEl.textContent = `${n}`;

    const chevron = createChevronIcon();
    chevron.classList.remove("pi-files-section-head__chevron");
    chevron.classList.add("pi-files-folder__chevron");

    header.append(iconEl, nameEl, countEl, chevron);

    const content = document.createElement("div");
    content.className = "pi-files-folder__content";
    folder.files.forEach((file) => {
      content.appendChild(createFileItem(file));
    });

    const applyCollapsed = (collapsed: boolean): void => {
      header.setAttribute("aria-expanded", collapsed ? "false" : "true");
      content.hidden = collapsed;
    };

    applyCollapsed(collapsedSections.has(collapseKey));

    header.addEventListener("click", () => {
      const isCollapsed = collapsedSections.has(collapseKey);
      if (isCollapsed) {
        collapsedSections.delete(collapseKey);
      } else {
        collapsedSections.add(collapseKey);
      }
      applyCollapsed(!isCollapsed);
    });

    group.append(header, content);
    return group;
  };

  const renderListView = (): void => {
    const files = allFiles;

    const connectState = resolveFilesDialogConnectFolderButtonState(backendStatus);
    connectFolderButton.hidden = connectState.hidden;
    connectFolderButton.disabled = connectState.disabled;
    connectFolderButton.title = connectState.title;
    connectFolderButton.setAttribute("aria-label", connectState.label);

    if (connectFolderButton.lastChild) {
      connectFolderButton.lastChild.textContent = ` ${connectState.label}`;
    }

    if (files.length < 5 && filterText.length > 0) {
      filterText = "";
    }

    if (files.length >= 5) {
      filterWrap.hidden = false;
      if (filterInput.value !== filterText) {
        filterInput.value = filterText;
      }
    } else {
      filterWrap.hidden = true;
      if (filterInput.value.length > 0) {
        filterInput.value = "";
      }
    }

    const sections = buildFilesDialogSections({
      files,
      filterText,
      backendStatus,
    });

    const hasNonBuiltInFiles = files.some((file) => !isFilesDialogBuiltInDoc(file));
    const hasFilterQuery = normalizeFilesDialogFilterText(filterText).length > 0;

    sectionsHost.replaceChildren();

    if (!hasNonBuiltInFiles && !hasFilterQuery) {
      sectionsHost.appendChild(createEmptyState(() => {
        hiddenInput.click();
      }));
    }

    if (sections.length === 0) {
      if (hasFilterQuery) {
        const empty = document.createElement("div");
        empty.className = "pi-files-empty";

        const title = document.createElement("div");
        title.className = "pi-files-empty__title";
        title.textContent = "No matching files";

        const description = document.createElement("p");
        description.className = "pi-files-empty__desc";
        description.textContent = "Try a different filter term.";

        empty.append(title, description);
        sectionsHost.appendChild(empty);
      }
    } else {
      sections.forEach((section) => {
        const sectionGroup = document.createElement("div");
        sectionGroup.className = "pi-files-section-group";

        const sectionHead = document.createElement("button");
        sectionHead.type = "button";
        sectionHead.className = "pi-files-section-head";

        const sectionLabel = document.createElement("span");
        sectionLabel.className = "pi-files-section-head__label";
        sectionLabel.textContent = section.label;

        const sectionCount = document.createElement("span");
        sectionCount.className = "pi-files-section-head__count";
        sectionCount.textContent = String(sectionTotalCount(section));

        sectionHead.append(sectionLabel, sectionCount, createChevronIcon());

        const sectionList = document.createElement("div");
        sectionList.className = "pi-files-section-list";

        // Root-level files
        section.files.forEach((file) => {
          sectionList.appendChild(createFileItem(file));
        });

        // Folder groups
        section.folders.forEach((folder) => {
          sectionList.appendChild(createFolderGroup(folder, `${section.key}:${folder.name}`));
        });

        const applyCollapsedState = (collapsed: boolean): void => {
          sectionHead.setAttribute("aria-expanded", collapsed ? "false" : "true");
          sectionList.hidden = collapsed;
        };

        applyCollapsedState(collapsedSections.has(section.key));

        sectionHead.addEventListener("click", () => {
          const currentlyCollapsed = collapsedSections.has(section.key);
          if (currentlyCollapsed) {
            collapsedSections.delete(section.key);
          } else {
            collapsedSections.add(section.key);
          }

          applyCollapsedState(!currentlyCollapsed);
        });

        sectionGroup.append(sectionHead, sectionList);
        sectionsHost.appendChild(sectionGroup);
      });
    }

    const totalSize = files.reduce((sum, file) => sum + file.size, 0);

    footer.textContent = buildFilesDialogStatusMessage({
      totalCount: files.length,
      totalSizeBytes: totalSize,
      backendLabel: backendStatus?.label ?? "Storage unavailable",
      nativeDirectoryName: backendStatus?.nativeConnected ? backendStatus.nativeDirectoryName ?? null : null,
    });
  };

  const onWorkspaceChanged: EventListener = () => {
    void (async () => {
      try {
        await refreshWorkspaceState();
        renderListView();

        if (currentView !== "detail" || !detailFileRef) {
          return;
        }

        const file = findFileByRef(detailFileRef);
        if (!file) {
          showToast("That file is no longer available.");
          showListView();
          return;
        }

        await renderDetailView(detailFileRef);
      } catch (error: unknown) {
        showToast(`Could not refresh files: ${getErrorMessage(error)}`);
      }
    })();
  };

  dialog.addCleanup(() => {
    document.removeEventListener(FILES_WORKSPACE_CHANGED_EVENT, onWorkspaceChanged);
    revokePreviewObjectUrl();
  });

  uploadButton.addEventListener("click", () => {
    hiddenInput.click();
  });

  connectFolderButton.addEventListener("click", () => {
    if (connectFolderButton.disabled) {
      return;
    }

    void workspace.connectNativeDirectory({
      audit: DIALOG_AUDIT_CONTEXT,
    }).catch((error: unknown) => {
      showToast(`Connect folder failed: ${getErrorMessage(error)}`);
    });
  });

  hiddenInput.addEventListener("change", () => {
    const selected = hiddenInput.files;
    if (!selected || selected.length === 0) {
      return;
    }

    const files = Array.from(selected);
    hiddenInput.value = "";

    void workspace.importFiles(files, {
      audit: DIALOG_AUDIT_CONTEXT,
    })
      .then((count) => {
        showToast(`Imported ${count} file${count === 1 ? "" : "s"}.`);
      })
      .catch((error: unknown) => {
        showToast(`Upload failed: ${getErrorMessage(error)}`);
      });
  });

  detailBackButton.addEventListener("click", () => {
    showListView();
  });

  document.addEventListener(FILES_WORKSPACE_CHANGED_EVENT, onWorkspaceChanged);

  dialog.mount();

  try {
    await refreshWorkspaceState();
    renderListView();
    setView("list");
  } catch (error: unknown) {
    showToast(`Could not load files: ${getErrorMessage(error)}`);
    showListView();
    renderListView();
  }
}
