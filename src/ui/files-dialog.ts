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
import { FILES_WORKSPACE_OVERLAY_ID } from "./overlay-ids.js";
import {
  buildFilesDialogFilterOptions,
  countBuiltInDocs,
  fileMatchesFilesDialogFilter,
  isFilesDialogFilterSelectable,
  parseFilesDialogFilterValue,
  type FilesDialogFilterOption,
  type FilesDialogFilterValue,
} from "./files-dialog-filtering.js";
import { buildFilesDialogStatusMessage } from "./files-dialog-status.js";
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

function makeButton(label: string, className: string): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = className;
  button.textContent = label;
  return button;
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
    cardClassName: "pi-welcome-card pi-overlay-card pi-files-dialog",
  });

  const closeOverlay = dialog.close;

  const { header, subtitle } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close files",
    title: "Files",
    subtitle: "",
  });

  if (!subtitle) {
    throw new Error("Files overlay subtitle is required.");
  }

  const controls = document.createElement("div");
  controls.className = "pi-files-dialog__controls";

  const enableButton = makeButton("Enable workspace write access", "pi-overlay-btn pi-overlay-btn--ghost");
  const uploadButton = makeButton("Upload", "pi-overlay-btn pi-overlay-btn--ghost");
  const newFileButton = makeButton("New text file", "pi-overlay-btn pi-overlay-btn--ghost");
  const nativeButton = makeButton("Select folder", "pi-overlay-btn pi-overlay-btn--ghost");
  const disconnectNativeButton = makeButton("Use sandbox workspace", "pi-overlay-btn pi-overlay-btn--ghost");

  const hiddenInput = document.createElement("input");
  hiddenInput.type = "file";
  hiddenInput.multiple = true;
  hiddenInput.className = "pi-files-dialog__hidden-input";

  const statusLine = document.createElement("div");
  statusLine.className = "pi-files-dialog__status";

  const helperLine = document.createElement("div");
  helperLine.className = "pi-files-dialog__helper";
  helperLine.textContent = "Built-in docs are always available. Enable write access for assistant file management.";

  const filters = document.createElement("div");
  filters.className = "pi-files-dialog__filters";

  const filterLabel = document.createElement("label");
  filterLabel.className = "pi-files-dialog__filter-label";
  filterLabel.textContent = "Filter";

  const filterSelect = document.createElement("select");
  filterSelect.className = "pi-files-dialog__filter-select";
  filterLabel.appendChild(filterSelect);

  const quickFilters = document.createElement("div");
  quickFilters.className = "pi-files-dialog__quick-filters";

  const quickAllButton = makeButton("All", "pi-files-dialog__quick-filter");
  const quickBuiltinButton = makeButton("Built-in docs", "pi-files-dialog__quick-filter");
  const quickCurrentButton = makeButton("Current workbook", "pi-files-dialog__quick-filter");

  quickAllButton.dataset.filter = "all";
  quickBuiltinButton.dataset.filter = "builtin";
  quickCurrentButton.dataset.filter = "current";

  quickFilters.append(quickAllButton, quickBuiltinButton, quickCurrentButton);
  filters.append(filterLabel, quickFilters);

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
    enableButton,
    uploadButton,
    newFileButton,
    nativeButton,
    disconnectNativeButton,
  );

  dialog.card.append(
    header,
    controls,
    hiddenInput,
    statusLine,
    helperLine,
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
    viewerNote.textContent = "Preview not available for this binary file. Use Download to inspect it locally.";

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

    filterSelect.replaceChildren();
    for (const option of filterOptions) {
      const optionElement = document.createElement("option");
      optionElement.value = option.value;
      optionElement.textContent = option.label;
      optionElement.disabled = option.disabled === true;
      optionElement.selected = option.value === selectedFilter;
      filterSelect.appendChild(optionElement);
    }
    filterSelect.disabled = files.length === 0;
    syncQuickFilterButtons(filterOptions);

    const filteredFiles = files.filter((file) => fileMatchesFilesDialogFilter({
      file,
      filter: selectedFilter,
      currentWorkbookId,
    }));

    const activeFilterLabel =
      filterOptions.find((option) => option.value === selectedFilter)?.label
      ?? "All files";

    const workspaceFilesCount = files.length - builtinDocsCount;

    enableButton.hidden = true;
    uploadButton.disabled = false;
    newFileButton.disabled = false;
    nativeButton.disabled = !backend.nativeSupported;
    nativeButton.hidden = !backend.nativeSupported;
    disconnectNativeButton.hidden = backend.kind !== "native-directory";

    setStatus(buildFilesDialogStatusMessage({
      filesExperimentEnabled: true,
      totalCount: files.length,
      filteredCount: filteredFiles.length,
      selectedFilter,
      activeFilterLabel,
      builtinDocsCount,
      workspaceFilesCount,
    }));

    list.replaceChildren();

    if (files.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-files-dialog__empty";
      empty.textContent = "No files yet. Upload documents or create a text file.";
      list.appendChild(empty);
    } else if (filteredFiles.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-files-dialog__empty";
      empty.textContent = "No files match the selected filter.";
      list.appendChild(empty);
    } else {
      for (const file of filteredFiles) {
        const row = document.createElement("div");
        row.className = "pi-files-dialog__row";

        const info = document.createElement("div");
        info.className = "pi-files-dialog__info";

        const nameRow = document.createElement("div");
        nameRow.className = "pi-files-dialog__name-row";

        const name = document.createElement("div");
        name.className = "pi-files-dialog__name";
        name.textContent = file.path;

        nameRow.appendChild(name);

        if (file.sourceKind === "builtin-doc") {
          const sourceBadge = document.createElement("span");
          sourceBadge.className = "pi-files-dialog__source-badge";
          sourceBadge.textContent = "Built-in";
          nameRow.appendChild(sourceBadge);
        }

        if (file.workbookTag) {
          const workbookTag = document.createElement("span");
          workbookTag.className = "pi-files-dialog__workbook-tag";
          workbookTag.textContent = file.workbookTag.workbookLabel;
          nameRow.appendChild(workbookTag);
        }

        const meta = document.createElement("div");
        meta.className = "pi-files-dialog__meta";
        meta.textContent = `${formatBytes(file.size)} · ${file.kind} · ${formatRelativeDate(file.modifiedAt)}`;

        info.append(nameRow, meta);

        const actions = document.createElement("div");
        actions.className = "pi-files-dialog__actions";

        const openButton = makeButton("Open", "pi-files-dialog__row-btn");
        openButton.addEventListener("click", () => {
          void openViewer(file);
        });

        const downloadButton = makeButton("Download", "pi-files-dialog__row-btn");
        downloadButton.addEventListener("click", () => {
          void workspace.downloadFile(file.path).catch((error: unknown) => {
            showToast(`Download failed: ${getErrorMessage(error)}`);
          });
        });

        const renameButton = makeButton("Rename", "pi-files-dialog__row-btn");
        renameButton.disabled = file.readOnly;
        if (file.readOnly) {
          renameButton.title = "Built-in docs are read-only.";
        }
        renameButton.addEventListener("click", () => {
          const nextName = window.prompt("Rename file", file.path);
          if (!nextName) return;

          void workspace.renameFile(file.path, nextName, {
            audit: DIALOG_AUDIT_CONTEXT,
          }).catch((error: unknown) => {
            showToast(`Rename failed: ${getErrorMessage(error)}`);
          });
        });

        const deleteButton = makeButton("Delete", "pi-files-dialog__row-btn pi-files-dialog__row-btn--danger");
        deleteButton.disabled = file.readOnly;
        if (file.readOnly) {
          deleteButton.title = "Built-in docs are read-only.";
        }
        deleteButton.addEventListener("click", () => {
          const ok = window.confirm(`Delete '${file.path}'?`);
          if (!ok) return;

          void workspace.deleteFile(file.path, {
            audit: DIALOG_AUDIT_CONTEXT,
          }).catch((error: unknown) => {
            showToast(`Delete failed: ${getErrorMessage(error)}`);
          });
        });

        actions.append(openButton, downloadButton, renameButton, deleteButton);
        row.append(info, actions);
        list.appendChild(row);
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

  filterSelect.addEventListener("change", () => {
    selectedFilter = parseFilesDialogFilterValue(filterSelect.value);
    void renderList();
  });

  for (const quickFilter of quickFilterButtons) {
    quickFilter.button.addEventListener("click", () => {
      selectedFilter = quickFilter.value;
      filterSelect.value = quickFilter.value;
      void renderList();
    });
  }

  newFileButton.addEventListener("click", () => {
    const path = window.prompt("New text file path", "notes.md");
    if (!path) return;

    void workspace.writeTextFile(path, "", undefined, {
      audit: DIALOG_AUDIT_CONTEXT,
    })
      .then(() => {
        showToast(`Created ${path}.`);
      })
      .catch((error: unknown) => {
        showToast(`Create failed: ${getErrorMessage(error)}`);
      });
  });

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

  disconnectNativeButton.addEventListener("click", () => {
    void workspace.disconnectNativeDirectory({
      audit: DIALOG_AUDIT_CONTEXT,
    })
      .then(() => {
        showToast("Switched to sandboxed workspace.");
      })
      .catch((error: unknown) => {
        showToast(`Could not switch workspace: ${getErrorMessage(error)}`);
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
