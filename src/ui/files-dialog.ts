/**
 * Files workspace dialog.
 */

import { base64ToBytes } from "../files/encoding.js";
import { formatBytes } from "../files/mime.js";
import {
  FILES_WORKSPACE_CHANGED_EVENT,
  type WorkspaceFileEntry,
} from "../files/types.js";
import { type FilesWorkspaceAuditContext, getFilesWorkspace } from "../files/workspace.js";
import { getErrorMessage } from "../utils/errors.js";
import { formatWorkbookLabel, getWorkbookContext } from "../workbook/context.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { requestConfirmationDialog } from "./confirm-dialog.js";
import { FILES_WORKSPACE_OVERLAY_ID } from "./overlay-ids.js";
import { requestTextInputDialog } from "./text-input-dialog.js";
import {
  buildFilesDialogFilterOptions,
  countBuiltInDocs,
  fileMatchesFilesDialogFilter,
  isFilesDialogFilterSelectable,
  type FilesDialogFilterOption,
  type FilesDialogFilterValue,
} from "./files-dialog-filtering.js";
import { buildFilesDialogStatusMessage } from "./files-dialog-status.js";
import { buildFilesDialogTree, type FilesDialogFolderNode } from "./files-dialog-tree.js";
import { showToast } from "./toast.js";

const OVERLAY_ID = FILES_WORKSPACE_OVERLAY_ID;

const DIALOG_AUDIT_CONTEXT: FilesWorkspaceAuditContext = {
  actor: "user",
  source: "files-dialog",
};

function formatRelativeDate(timestamp: number): string {
  const now = Date.now();
  const diff = now - timestamp;

  if (diff < 60_000) return "just now";
  if (diff < 3_600_000) return `${Math.round(diff / 60_000)}m ago`;
  if (diff < 86_400_000) return `${Math.round(diff / 3_600_000)}h ago`;
  if (diff < 604_800_000) return `${Math.round(diff / 86_400_000)}d ago`;
  return new Date(timestamp).toLocaleDateString();
}

function formatFileCountLabel(count: number): string {
  return `${count} file${count === 1 ? "" : "s"}`;
}

const WORKSPACE_ROOT_COLLAPSE_KEY = "workspace-root";

function getFolderCollapseKey(folderPath: string): string {
  return `path:${folderPath}`;
}

function makeButton(label: string, className: string): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = className;
  button.textContent = label;
  return button;
}

function getFileExtension(fileName: string): string | null {
  const trimmed = fileName.trim();
  const lastDot = trimmed.lastIndexOf(".");
  if (lastDot <= 0 || lastDot === trimmed.length - 1) {
    return null;
  }

  return trimmed.slice(lastDot + 1).toLowerCase();
}

function getBinaryPreviewNote(file: WorkspaceFileEntry): string {
  const extension = getFileExtension(file.name);
  if (!extension) {
    return "Preview not available for this binary file. Use Download to inspect it locally.";
  }

  return `Preview not available for .${extension} files. Use Download to inspect locally.`;
}

function resolveRenameDestinationPath(currentPath: string, inputPath: string): string {
  const normalizedInput = inputPath.trim().replaceAll("\\", "/");
  if (normalizedInput.length === 0) {
    return currentPath;
  }

  if (normalizedInput.includes("/")) {
    return normalizedInput;
  }

  const lastSlash = currentPath.lastIndexOf("/");
  if (lastSlash < 0) {
    return normalizedInput;
  }

  return `${currentPath.slice(0, lastSlash + 1)}${normalizedInput}`;
}

function isImageMimeType(mimeType: string): boolean {
  return mimeType.toLowerCase().startsWith("image/");
}

function isPdfMimeType(mimeType: string): boolean {
  return mimeType.trim().toLowerCase() === "application/pdf";
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
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

  const { header, subtitle } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close files",
    title: "Files",
    subtitle: "Storage: loading…",
  });

  if (!subtitle) {
    throw new Error("Files overlay subtitle is required.");
  }

  const controls = document.createElement("div");
  controls.className = "pi-files-dialog__controls";

  const uploadButton = makeButton("Upload", "pi-overlay-btn pi-overlay-btn--ghost");
  const nativeButton = makeButton("Select folder", "pi-overlay-btn pi-overlay-btn--ghost");

  const hiddenInput = document.createElement("input");
  hiddenInput.type = "file";
  hiddenInput.multiple = true;
  hiddenInput.className = "pi-files-dialog__hidden-input";

  const statusLine = document.createElement("div");
  statusLine.className = "pi-files-dialog__status";

  const filters = document.createElement("div");
  filters.className = "pi-files-dialog__filters";

  const quickFilters = document.createElement("div");
  quickFilters.className = "pi-files-dialog__quick-filters";

  const quickAllButton = makeButton("All", "pi-files-dialog__quick-filter");
  const quickBuiltinButton = makeButton("Built-in docs", "pi-files-dialog__quick-filter");
  const quickCurrentButton = makeButton("Current workbook", "pi-files-dialog__quick-filter");

  quickAllButton.dataset.filter = "all";
  quickBuiltinButton.dataset.filter = "builtin";
  quickCurrentButton.dataset.filter = "current";

  quickFilters.append(quickAllButton, quickBuiltinButton, quickCurrentButton);
  filters.append(quickFilters);

  const list = document.createElement("div");
  list.className = "pi-files-dialog__list";

  const viewer = document.createElement("div");
  viewer.className = "pi-files-dialog__viewer";
  viewer.hidden = true;

  const viewerHeader = document.createElement("div");
  viewerHeader.className = "pi-files-dialog__viewer-header";

  const viewerTitle = document.createElement("div");
  viewerTitle.className = "pi-files-dialog__viewer-title";

  const viewerActions = document.createElement("div");
  viewerActions.className = "pi-files-dialog__viewer-actions";

  const saveButton = makeButton("Save", "pi-overlay-btn pi-overlay-btn--primary");
  const closeViewerButton = makeButton("Close", "pi-overlay-btn pi-overlay-btn--ghost");

  viewerActions.append(saveButton, closeViewerButton);
  viewerHeader.append(viewerTitle, viewerActions);

  const viewerNote = document.createElement("div");
  viewerNote.className = "pi-files-dialog__viewer-note";

  const viewerTextarea = document.createElement("textarea");
  viewerTextarea.className = "pi-files-dialog__textarea";
  viewerTextarea.spellcheck = false;

  const viewerPreview = document.createElement("div");
  viewerPreview.className = "pi-files-dialog__preview";
  viewerPreview.hidden = true;

  viewer.append(viewerHeader, viewerNote, viewerTextarea, viewerPreview);

  controls.append(
    uploadButton,
    nativeButton,
  );

  dialog.card.append(
    header,
    controls,
    hiddenInput,
    statusLine,
    filters,
    list,
    viewer,
  );

  let activeViewerPath: string | null = null;
  let activeViewerReadOnly = false;
  let viewerTruncated = false;
  let activeObjectUrl: string | null = null;
  let selectedFilter: FilesDialogFilterValue = "all";
  let currentWorkbookId: string | null = null;
  let currentWorkbookLabel: string | null = null;
  const collapsedFolderPaths = new Set<string>();

  const quickFilterButtons: Array<{
    value: FilesDialogFilterValue;
    button: HTMLButtonElement;
  }> = [
    { value: "all", button: quickAllButton },
    { value: "builtin", button: quickBuiltinButton },
    { value: "current", button: quickCurrentButton },
  ];

  const revokeObjectUrl = () => {
    if (!activeObjectUrl) return;
    URL.revokeObjectURL(activeObjectUrl);
    activeObjectUrl = null;
  };

  const setStatus = (message: string) => {
    statusLine.textContent = message;
  };

  const syncQuickFilterButtons = (options: FilesDialogFilterOption[]) => {
    for (const quickFilter of quickFilterButtons) {
      const option = options.find((candidate) => candidate.value === quickFilter.value);
      const isSelectable = option !== undefined && option.disabled !== true;
      quickFilter.button.disabled = !isSelectable;

      const isActive = selectedFilter === quickFilter.value;
      quickFilter.button.classList.toggle("is-active", isActive);
      quickFilter.button.setAttribute("aria-pressed", isActive ? "true" : "false");
    }
  };

  const setViewerMode = (mode: "hidden" | "text" | "preview") => {
    viewer.hidden = mode === "hidden";
    viewerTextarea.hidden = mode !== "text";
    viewerPreview.hidden = mode !== "preview";
  };

  const clearViewer = () => {
    activeViewerPath = null;
    activeViewerReadOnly = false;
    viewerTruncated = false;
    revokeObjectUrl();

    viewerTitle.textContent = "";
    viewerNote.textContent = "";
    viewerTextarea.value = "";
    viewerTextarea.disabled = false;
    viewerPreview.replaceChildren();

    saveButton.hidden = true;
    saveButton.disabled = true;

    setViewerMode("hidden");
  };

  const openTextViewer = async (entry: WorkspaceFileEntry) => {
    const result = await workspace.readFile(entry.path, {
      mode: "text",
      maxChars: 1_000_000,
      audit: DIALOG_AUDIT_CONTEXT,
    });

    activeViewerPath = entry.path;
    activeViewerReadOnly = entry.readOnly;
    viewerTruncated = result.truncated === true;

    viewerTitle.textContent = entry.path;
    viewerTextarea.value = result.text ?? "";
    viewerTextarea.disabled = viewerTruncated || entry.readOnly;

    if (entry.readOnly) {
      viewerNote.textContent = "Built-in documentation file (read-only).";
      saveButton.hidden = true;
      saveButton.disabled = true;
    } else if (viewerTruncated) {
      viewerNote.textContent = "This file is too large to edit inline safely (preview truncated to 1,000,000 chars).";
      saveButton.hidden = false;
      saveButton.disabled = true;
    } else {
      viewerNote.textContent = "Editable text file.";
      saveButton.hidden = false;
      saveButton.disabled = false;
    }

    setViewerMode("text");
  };

  const openImagePreview = async (entry: WorkspaceFileEntry) => {
    const result = await workspace.readFile(entry.path, {
      mode: "base64",
      maxChars: 8_000_000,
      audit: DIALOG_AUDIT_CONTEXT,
    });

    viewerTitle.textContent = entry.path;
    saveButton.hidden = true;
    saveButton.disabled = true;

    if (!result.base64 || result.truncated) {
      viewerNote.textContent = "Preview unavailable: image is too large for inline preview.";
      viewerPreview.replaceChildren();
      setViewerMode("preview");
      return;
    }

    const bytes = base64ToBytes(result.base64);
    const blob = new Blob([toArrayBuffer(bytes)], { type: entry.mimeType });
    activeObjectUrl = URL.createObjectURL(blob);

    const image = document.createElement("img");
    image.className = "pi-files-dialog__preview-image";
    image.src = activeObjectUrl;
    image.alt = entry.path;

    viewerNote.textContent = `Image preview (${formatBytes(entry.size)}).`;
    viewerPreview.replaceChildren(image);
    setViewerMode("preview");
  };

  const openPdfPreview = async (entry: WorkspaceFileEntry) => {
    const result = await workspace.readFile(entry.path, {
      mode: "base64",
      maxChars: 16_000_000,
      audit: DIALOG_AUDIT_CONTEXT,
    });

    viewerTitle.textContent = entry.path;
    saveButton.hidden = true;
    saveButton.disabled = true;

    if (!result.base64 || result.truncated) {
      viewerNote.textContent = "Preview unavailable: PDF is too large for inline preview.";
      viewerPreview.replaceChildren();
      setViewerMode("preview");
      return;
    }

    const bytes = base64ToBytes(result.base64);
    const blob = new Blob([toArrayBuffer(bytes)], { type: "application/pdf" });
    activeObjectUrl = URL.createObjectURL(blob);

    const frame = document.createElement("iframe");
    frame.className = "pi-files-dialog__preview-frame";
    frame.src = activeObjectUrl;
    frame.title = entry.path;

    viewerNote.textContent = `PDF preview (${formatBytes(entry.size)}).`;
    viewerPreview.replaceChildren(frame);
    setViewerMode("preview");
  };

  const openBinaryPlaceholder = (entry: WorkspaceFileEntry) => {
    viewerTitle.textContent = entry.path;
    viewerNote.textContent = getBinaryPreviewNote(entry);

    const message = document.createElement("div");
    message.className = "pi-files-dialog__preview-empty";
    message.textContent = `${entry.mimeType} · ${formatBytes(entry.size)}`;

    viewerPreview.replaceChildren(message);
    saveButton.hidden = true;
    saveButton.disabled = true;
    setViewerMode("preview");
  };

  const openViewer = async (entry: WorkspaceFileEntry) => {
    try {
      revokeObjectUrl();
      activeViewerPath = null;
      activeViewerReadOnly = false;
      viewerTruncated = false;
      viewerPreview.replaceChildren();

      if (entry.kind === "text") {
        await openTextViewer(entry);
        return;
      }

      if (isImageMimeType(entry.mimeType)) {
        await openImagePreview(entry);
        return;
      }

      if (isPdfMimeType(entry.mimeType)) {
        await openPdfPreview(entry);
        return;
      }

      openBinaryPlaceholder(entry);
    } catch (error: unknown) {
      activeViewerPath = null;
      activeViewerReadOnly = false;
      viewerTruncated = false;
      viewerTitle.textContent = entry.path;
      viewerNote.textContent = `Preview unavailable: ${getErrorMessage(error)}`;
      viewerTextarea.value = "";
      viewerTextarea.disabled = true;
      saveButton.hidden = true;
      saveButton.disabled = true;

      const message = document.createElement("div");
      message.className = "pi-files-dialog__preview-empty";
      message.textContent = "Try downloading the file and opening it locally.";
      viewerPreview.replaceChildren(message);
      setViewerMode("preview");
    }
  };

  const createFileRow = (file: WorkspaceFileEntry): HTMLElement => {
    const row = document.createElement("div");
    row.className = "pi-files-row";
    row.tabIndex = 0;
    row.setAttribute("role", "listitem");
    row.setAttribute("aria-label", `${file.path} — press Enter to open`);

    const fileIcon = file.kind === "text" ? "📄" : isImageMimeType(file.mimeType) ? "🖼" : "📎";

    const info = document.createElement("div");
    info.className = "pi-files-row__info";

    const nameRow = document.createElement("div");
    nameRow.className = "pi-files-row__name-row";

    const iconEl = document.createElement("span");
    iconEl.className = "pi-files-row__icon";
    iconEl.textContent = fileIcon;

    const name = document.createElement("div");
    name.className = "pi-files-row__name";
    name.textContent = file.path;

    nameRow.append(iconEl, name);

    if (file.sourceKind === "builtin-doc") {
      const sourceBadge = document.createElement("span");
      sourceBadge.className = "pi-files-row__badge pi-files-row__badge--info";
      sourceBadge.textContent = "Built-in";
      nameRow.appendChild(sourceBadge);
    }

    if (file.workbookTag) {
      const workbookTag = document.createElement("span");
      workbookTag.className = "pi-files-row__badge pi-files-row__badge--ok";
      workbookTag.textContent = file.workbookTag.workbookLabel;
      nameRow.appendChild(workbookTag);
    }

    const meta = document.createElement("div");
    meta.className = "pi-files-row__meta";
    const metaText = document.createElement("span");
    metaText.textContent = `${formatBytes(file.size)} · ${file.kind} · ${formatRelativeDate(file.modifiedAt)}`;
    meta.appendChild(metaText);

    info.append(nameRow, meta);

    // Click or keyboard activate row to open viewer
    const activateRow = (event: Event) => {
      if ((event.target as HTMLElement).closest(".pi-files-row__overflow")) return;
      if ((event.target as HTMLElement).closest(".pi-files-overflow-menu")) return;
      void openViewer(file);
    };
    row.addEventListener("click", activateRow);
    row.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        activateRow(event);
      }
    });

    // Overflow menu (⋯)
    const overflowBtn = document.createElement("button");
    overflowBtn.type = "button";
    overflowBtn.className = "pi-files-row__overflow";
    overflowBtn.textContent = "⋯";
    overflowBtn.title = "File actions";

    overflowBtn.addEventListener("click", (event) => {
      event.stopPropagation();

      // Close any existing overflow menu
      const existing = document.querySelector(".pi-files-overflow-menu");
      if (existing) {
        existing.remove();
        return;
      }

      const menu = document.createElement("div");
      menu.className = "pi-files-overflow-menu";

      const addMenuItem = (label: string, handler: () => void, tone?: "danger") => {
        const item = document.createElement("button");
        item.type = "button";
        item.className = `pi-files-overflow-menu__item${tone === "danger" ? " pi-files-overflow-menu__item--danger" : ""}`;
        item.textContent = label;
        item.addEventListener("click", () => {
          menu.remove();
          handler();
        });
        menu.appendChild(item);
      };

      addMenuItem("Download", () => {
        void workspace.downloadFile(file.path).catch((error: unknown) => {
          showToast(`Download failed: ${getErrorMessage(error)}`);
        });
      });

      if (!file.readOnly) {
        addMenuItem("Rename", () => {
          void (async () => {
            const nextPathInput = await requestTextInputDialog({
              title: "Rename file",
              message: file.path,
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
            });

            showToast(`Renamed to ${nextPath}.`);
          })().catch((error: unknown) => {
            showToast(`Rename failed: ${getErrorMessage(error)}`);
          });
        });

        const sep = document.createElement("div");
        sep.className = "pi-files-overflow-menu__separator";
        menu.appendChild(sep);

        addMenuItem("Delete", () => {
          void (async () => {
            const ok = await requestConfirmationDialog({
              title: "Delete file?",
              message: file.path,
              confirmLabel: "Delete",
              cancelLabel: "Cancel",
              confirmButtonTone: "danger",
              restoreFocusOnClose: false,
            });
            if (!ok) return;
            await workspace.deleteFile(file.path, {
              audit: DIALOG_AUDIT_CONTEXT,
            }).catch((error: unknown) => {
              showToast(`Delete failed: ${getErrorMessage(error)}`);
            });
          })();
        }, "danger");
      }

      meta.appendChild(menu);

      const closeOnOutsideClick = (e: MouseEvent) => {
        if (!menu.contains(e.target as Node)) {
          menu.remove();
          document.removeEventListener("click", closeOnOutsideClick, true);
        }
      };
      // Delay listener to avoid catching the current click
      requestAnimationFrame(() => {
        document.addEventListener("click", closeOnOutsideClick, true);
      });
    });

    meta.appendChild(overflowBtn);
    row.appendChild(info);

    return row;
  };

  const appendFolderNode = (
    container: HTMLElement,
    folder: FilesDialogFolderNode,
    collapseKey: string,
  ): void => {
    const section = document.createElement("section");
    section.className = "pi-files-section";

    const header = document.createElement("button");
    header.type = "button";
    header.className = "pi-files-section__header";

    const label = document.createElement("span");
    label.className = "pi-files-section__label";
    label.textContent = folder.folderName.toUpperCase();

    const count = document.createElement("span");
    count.className = "pi-files-section__count";
    count.textContent = formatFileCountLabel(folder.totalFileCount);

    header.append(label, count);

    const body = document.createElement("div");
    body.className = "pi-files-section__body";

    const applyCollapsedState = (isCollapsed: boolean): void => {
      section.classList.toggle("is-collapsed", isCollapsed);
      header.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
      body.hidden = isCollapsed;
    };

    let isCollapsed = collapsedFolderPaths.has(collapseKey);
    applyCollapsedState(isCollapsed);

    for (const file of folder.files) {
      body.appendChild(createFileRow(file));
    }

    for (const childFolder of folder.children) {
      appendFolderNode(body, childFolder, getFolderCollapseKey(childFolder.folderPath));
    }

    header.addEventListener("click", () => {
      isCollapsed = !isCollapsed;
      if (isCollapsed) {
        collapsedFolderPaths.add(collapseKey);
      } else {
        collapsedFolderPaths.delete(collapseKey);
      }

      applyCollapsedState(isCollapsed);
    });

    section.append(header, body);
    container.appendChild(section);
  };

  const renderList = async () => {
    const [backend, files, workbookContext] = await Promise.all([
      workspace.getBackendStatus(),
      workspace.listFiles(),
      getWorkbookContext().catch(() => null),
    ]);

    subtitle.textContent = `Storage: ${backend.label}${backend.nativeDirectoryName ? ` (${backend.nativeDirectoryName})` : ""}`;

    currentWorkbookId = workbookContext?.workbookId ?? null;
    currentWorkbookLabel = workbookContext && workbookContext.workbookId
      ? formatWorkbookLabel(workbookContext)
      : null;

    const builtinDocsCount = countBuiltInDocs(files);
    quickBuiltinButton.textContent = builtinDocsCount > 0
      ? `Built-in docs (${builtinDocsCount})`
      : "Built-in docs";

    const filterOptions = buildFilesDialogFilterOptions({
      files,
      currentWorkbookId,
      currentWorkbookLabel,
      builtinDocsCount,
    });

    if (!isFilesDialogFilterSelectable({
      filter: selectedFilter,
      options: filterOptions,
    })) {
      selectedFilter = "all";
    }

    syncQuickFilterButtons(filterOptions);

    const filteredFiles = files.filter((file) => fileMatchesFilesDialogFilter({
      file,
      filter: selectedFilter,
      currentWorkbookId,
    }));

    const activeFilterLabel =
      filterOptions.find((option) => option.value === selectedFilter)?.label
      ?? "All files";

    nativeButton.disabled = !backend.nativeSupported;
    nativeButton.hidden = !backend.nativeSupported;

    setStatus(buildFilesDialogStatusMessage({
      totalCount: files.length,
      filteredCount: filteredFiles.length,
      selectedFilter,
      activeFilterLabel,
    }));

    list.replaceChildren();

    if (files.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-files-dialog__empty";
      empty.textContent = "No files yet. Upload documents to get started.";
      list.appendChild(empty);
    } else if (filteredFiles.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-files-dialog__empty";
      empty.textContent = "No files match the selected filter.";
      list.appendChild(empty);
    } else {
      const tree = buildFilesDialogTree(filteredFiles);

      if (tree.rootFiles.length > 0) {
        appendFolderNode(list, {
          folderName: "workspace root",
          folderPath: "",
          files: tree.rootFiles,
          children: [],
          totalFileCount: tree.rootFiles.length,
        }, WORKSPACE_ROOT_COLLAPSE_KEY);
      }

      for (const folder of tree.folders) {
        appendFolderNode(list, folder, getFolderCollapseKey(folder.folderPath));
      }
    }

  };

  const onWorkspaceChanged: EventListener = () => {
    void renderList();
  };

  const cleanup = () => {
    document.removeEventListener(FILES_WORKSPACE_CHANGED_EVENT, onWorkspaceChanged);
    revokeObjectUrl();
  };

  dialog.addCleanup(cleanup);

  uploadButton.addEventListener("click", () => {
    hiddenInput.click();
  });

  hiddenInput.addEventListener("change", () => {
    const { files } = hiddenInput;
    if (!files || files.length === 0) return;

    const selectedFiles = Array.from(files);
    hiddenInput.value = "";

    void workspace.importFiles(selectedFiles, {
      audit: DIALOG_AUDIT_CONTEXT,
    })
      .then((count) => {
        showToast(`Imported ${count} file${count === 1 ? "" : "s"}.`);
      })
      .catch((error: unknown) => {
        showToast(`Upload failed: ${getErrorMessage(error)}`);
      });
  });

  for (const quickFilter of quickFilterButtons) {
    quickFilter.button.addEventListener("click", () => {
      selectedFilter = quickFilter.value;
      void renderList();
    });
  }

  nativeButton.addEventListener("click", () => {
    void workspace.connectNativeDirectory({
      audit: DIALOG_AUDIT_CONTEXT,
    })
      .then(() => {
        showToast("Connected local folder.");
      })
      .catch((error: unknown) => {
        showToast(`Could not connect folder: ${getErrorMessage(error)}`);
      });
  });

  saveButton.addEventListener("click", () => {
    if (!activeViewerPath || viewerTruncated || activeViewerReadOnly) return;

    const path = activeViewerPath;
    const nextContent = viewerTextarea.value;
    void workspace.writeTextFile(path, nextContent, undefined, {
      audit: DIALOG_AUDIT_CONTEXT,
    })
      .then(() => {
        showToast(`Saved ${path}.`);
      })
      .catch((error: unknown) => {
        showToast(`Save failed: ${getErrorMessage(error)}`);
      });
  });

  closeViewerButton.addEventListener("click", () => {
    clearViewer();
  });

  document.addEventListener(FILES_WORKSPACE_CHANGED_EVENT, onWorkspaceChanged);

  dialog.mount();
  await renderList();
}
