/**
 * Files workspace dialog.
 */

import { base64ToBytes } from "../files/encoding.js";
import { formatBytes } from "../files/mime.js";
import {
  FILES_WORKSPACE_CHANGED_EVENT,
  type FilesWorkspaceAuditEntry,
  type WorkspaceFileEntry,
} from "../files/types.js";
import { type FilesWorkspaceAuditContext, getFilesWorkspace } from "../files/workspace.js";
import { isExperimentalFeatureEnabled, setExperimentalFeatureEnabled } from "../experiments/flags.js";
import { getErrorMessage } from "../utils/errors.js";
import { showToast } from "./toast.js";

const OVERLAY_ID = "pi-files-workspace-overlay";

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

function describeAuditEntry(entry: FilesWorkspaceAuditEntry): string {
  switch (entry.action) {
    case "list":
      return "Listed workspace files";
    case "read":
      return entry.path ? `Read ${entry.path}` : "Read file";
    case "write":
      return entry.path ? `Wrote ${entry.path}` : "Wrote file";
    case "delete":
      return entry.path ? `Deleted ${entry.path}` : "Deleted file";
    case "rename":
      if (entry.fromPath && entry.toPath) {
        return `Renamed ${entry.fromPath} → ${entry.toPath}`;
      }
      return "Renamed file";
    case "import":
      return "Imported files";
    case "connect_native":
      return "Connected local folder";
    case "disconnect_native":
      return "Switched to sandbox workspace";
    case "clear_audit":
      return "Cleared audit trail";
  }
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
  const existing = document.getElementById(OVERLAY_ID);
  if (existing) {
    existing.remove();
    return;
  }

  const workspace = getFilesWorkspace();

  const overlay = document.createElement("div");
  overlay.id = OVERLAY_ID;
  overlay.className = "pi-welcome-overlay";

  const card = document.createElement("div");
  card.className = "pi-welcome-card pi-files-dialog";

  const title = document.createElement("h2");
  title.className = "pi-files-dialog__title";
  title.textContent = "Files workspace";

  const subtitle = document.createElement("p");
  subtitle.className = "pi-files-dialog__subtitle";

  const controls = document.createElement("div");
  controls.className = "pi-files-dialog__controls";

  const enableButton = makeButton("Enable assistant access", "pi-files-dialog__btn");
  const uploadButton = makeButton("Upload", "pi-files-dialog__btn");
  const newFileButton = makeButton("New text file", "pi-files-dialog__btn");
  const nativeButton = makeButton("Select folder", "pi-files-dialog__btn");
  const disconnectNativeButton = makeButton("Use sandbox workspace", "pi-files-dialog__btn");

  const hiddenInput = document.createElement("input");
  hiddenInput.type = "file";
  hiddenInput.multiple = true;
  hiddenInput.className = "pi-files-dialog__hidden-input";

  const statusLine = document.createElement("div");
  statusLine.className = "pi-files-dialog__status";

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

  const saveButton = makeButton("Save", "pi-files-dialog__btn pi-files-dialog__btn--primary");
  const closeViewerButton = makeButton("Close", "pi-files-dialog__btn");

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

  const audit = document.createElement("div");
  audit.className = "pi-files-dialog__audit";

  const auditHeader = document.createElement("div");
  auditHeader.className = "pi-files-dialog__audit-header";

  const auditTitle = document.createElement("div");
  auditTitle.className = "pi-files-dialog__audit-title";
  auditTitle.textContent = "Recent activity";

  const clearAuditButton = makeButton("Clear", "pi-files-dialog__row-btn");

  auditHeader.append(auditTitle, clearAuditButton);

  const auditList = document.createElement("div");
  auditList.className = "pi-files-dialog__audit-list";

  audit.append(auditHeader, auditList);

  const footer = document.createElement("div");
  footer.className = "pi-files-dialog__footer";
  const closeButton = makeButton("Close", "pi-files-dialog__btn");
  footer.appendChild(closeButton);

  controls.append(
    enableButton,
    uploadButton,
    newFileButton,
    nativeButton,
    disconnectNativeButton,
  );

  card.append(
    title,
    subtitle,
    controls,
    hiddenInput,
    statusLine,
    list,
    viewer,
    audit,
    footer,
  );

  overlay.appendChild(card);

  let activeViewerPath: string | null = null;
  let viewerTruncated = false;
  let activeObjectUrl: string | null = null;

  const revokeObjectUrl = () => {
    if (!activeObjectUrl) return;
    URL.revokeObjectURL(activeObjectUrl);
    activeObjectUrl = null;
  };

  const closeOverlay = () => {
    cleanup();
    revokeObjectUrl();
    overlay.remove();
  };

  const setStatus = (message: string) => {
    statusLine.textContent = message;
  };

  const setViewerMode = (mode: "hidden" | "text" | "preview") => {
    viewer.hidden = mode === "hidden";
    viewerTextarea.hidden = mode !== "text";
    viewerPreview.hidden = mode !== "preview";
  };

  const clearViewer = () => {
    activeViewerPath = null;
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
    viewerTruncated = result.truncated === true;

    viewerTitle.textContent = entry.path;
    viewerTextarea.value = result.text ?? "";
    viewerTextarea.disabled = viewerTruncated;

    if (viewerTruncated) {
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

  const renderAuditTrail = (entries: FilesWorkspaceAuditEntry[]) => {
    auditList.replaceChildren();

    if (entries.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-files-dialog__empty";
      empty.textContent = "No activity yet.";
      auditList.appendChild(empty);
      return;
    }

    for (const entry of entries.slice(0, 40)) {
      const row = document.createElement("div");
      row.className = "pi-files-dialog__audit-row";

      const meta = document.createElement("div");
      meta.className = "pi-files-dialog__audit-meta";

      const actorLabel = entry.actor === "assistant"
        ? "assistant"
        : entry.actor === "system"
          ? "system"
          : "user";

      meta.textContent = `${formatRelativeDate(entry.at)} · ${actorLabel} · ${entry.source}`;

      const body = document.createElement("div");
      body.className = "pi-files-dialog__audit-body";

      const workbookSuffix = entry.workbookLabel ? ` · ${entry.workbookLabel}` : "";
      body.textContent = `${describeAuditEntry(entry)}${workbookSuffix}`;

      row.append(meta, body);
      auditList.appendChild(row);
    }
  };

  const renderList = async () => {
    const [backend, files, auditEntries] = await Promise.all([
      workspace.getBackendStatus(),
      workspace.listFiles(),
      workspace.listAuditEntries(80),
    ]);

    subtitle.textContent = `Storage: ${backend.label}${backend.nativeDirectoryName ? ` (${backend.nativeDirectoryName})` : ""}`;

    const filesExperimentEnabled = isExperimentalFeatureEnabled("files_workspace");
    enableButton.hidden = filesExperimentEnabled;
    uploadButton.disabled = !filesExperimentEnabled;
    newFileButton.disabled = !filesExperimentEnabled;
    nativeButton.disabled = !filesExperimentEnabled || !backend.nativeSupported;
    nativeButton.hidden = !backend.nativeSupported;
    disconnectNativeButton.hidden = backend.kind !== "native-directory";

    if (!filesExperimentEnabled) {
      setStatus("Assistant access is disabled. Enable files-workspace to expose the tool.");
    } else {
      setStatus(`${files.length} file${files.length === 1 ? "" : "s"} available to the assistant.`);
    }

    list.replaceChildren();

    if (files.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-files-dialog__empty";
      empty.textContent = "No files yet. Upload documents or create a text file.";
      list.appendChild(empty);
    } else {
      for (const file of files) {
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

    renderAuditTrail(auditEntries);
  };

  const onWorkspaceChanged: EventListener = () => {
    void renderList();
  };

  const onEscape = (event: KeyboardEvent) => {
    if (event.key !== "Escape") return;
    if (event.target instanceof HTMLElement && (event.target.tagName === "TEXTAREA" || event.target.tagName === "INPUT")) {
      return;
    }

    event.preventDefault();
    closeOverlay();
  };

  const cleanup = () => {
    document.removeEventListener(FILES_WORKSPACE_CHANGED_EVENT, onWorkspaceChanged);
    document.removeEventListener("keydown", onEscape, true);
  };

  enableButton.addEventListener("click", () => {
    setExperimentalFeatureEnabled("files_workspace", true);
    void renderList();
    showToast("Enabled experimental files workspace.");
  });

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
    if (!activeViewerPath || viewerTruncated) return;

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

  clearAuditButton.addEventListener("click", () => {
    const ok = window.confirm("Clear files activity log?");
    if (!ok) return;

    void workspace.clearAuditTrail({
      audit: DIALOG_AUDIT_CONTEXT,
    })
      .then(() => {
        showToast("Cleared files activity log.");
      })
      .catch((error: unknown) => {
        showToast(`Could not clear activity log: ${getErrorMessage(error)}`);
      });
  });

  closeButton.addEventListener("click", closeOverlay);

  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      closeOverlay();
    }
  });

  document.addEventListener(FILES_WORKSPACE_CHANGED_EVENT, onWorkspaceChanged);
  document.addEventListener("keydown", onEscape, true);

  document.body.appendChild(overlay);
  await renderList();
}
