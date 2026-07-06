import { base64ToBytes } from "../files/encoding.js";
import { resolveSafeBlobUrlMimeType } from "../files/blob-url-safety.js";
import type { WorkspaceFileEntry, WorkspaceFileLocationKind, WorkspaceFileReadResult } from "../files/types.js";
import type {
  FilesWorkspaceAuditContext,
  WorkspaceMutationOptions,
  WorkspaceReadOptions,
} from "../files/workspace.js";
import { getErrorMessage } from "../utils/errors.js";
import { requestConfirmationDialog } from "./confirm-dialog.js";
import { isFilesDialogBuiltInDoc } from "./files-dialog-filtering.js";
import { resolveRenameDestinationPath } from "./files-dialog-paths.js";
import { requestTextInputDialog } from "./text-input-dialog.js";
import { showToast } from "./toast.js";
import { t } from "../language/index.js";

export interface FilesDialogDetailActionFileRef {
  path: string;
  locationKind: WorkspaceFileLocationKind;
}

export interface FilesDialogDetailActionsWorkspace {
  readFile(path: string, opts?: WorkspaceReadOptions): Promise<WorkspaceFileReadResult>;
  downloadFile(path: string, options?: { locationKind?: WorkspaceFileLocationKind }): Promise<void>;
  renameFile(oldPath: string, newPath: string, options?: WorkspaceMutationOptions): Promise<void>;
  deleteFile(path: string, options?: WorkspaceMutationOptions): Promise<void>;
}

export interface CreateFilesDialogDetailActionsOptions {
  file: WorkspaceFileEntry;
  fileRef: FilesDialogDetailActionFileRef;
  workspace: FilesDialogDetailActionsWorkspace;
  auditContext: FilesWorkspaceAuditContext;
  onAfterRename: (nextPath: string, locationKind: WorkspaceFileLocationKind) => Promise<void>;
  onAfterDelete: () => Promise<void>;
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
    // Best-effort — anchor click may also silently fail in WebView.
    // We can't verify whether it worked, so no error toast here.
  }

  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 60_000);
}

/**
 * Copy text content to clipboard — reliable in Office WebView where
 * window.open / blob URLs silently fail.
 */
async function copyTextToClipboard(text: string, fileName: string): Promise<void> {
  await navigator.clipboard.writeText(text);
  showToast(t("files-dialog-actions.toast.copied", { fileName }));
}

/**
 * Download via data URI — works in WebView where blob URLs may not.
 */
function downloadViaDataUri(text: string, fileName: string, mimeType: string): void {
  const dataUri = `data:${mimeType};charset=utf-8,${encodeURIComponent(text)}`;
  const anchor = document.createElement("a");
  anchor.href = dataUri;
  anchor.download = fileName;
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

async function openFileInBrowser(options: {
  file: WorkspaceFileEntry;
  fileRef: FilesDialogDetailActionFileRef;
  workspace: FilesDialogDetailActionsWorkspace;
  auditContext: FilesWorkspaceAuditContext;
}): Promise<void> {
  const pendingWindow = window.open("", "_blank");

  try {
    if (options.file.kind === "text") {
      const result = await options.workspace.readFile(options.file.path, {
        mode: "text",
        maxChars: 16_000_000,
        audit: options.auditContext,
        locationKind: options.fileRef.locationKind,
      });

      if (result.text === undefined || result.truncated) {
        throw new Error("File is too large to open in a browser tab.");
      }

      const blob = new Blob([result.text], {
        type: resolveSafeBlobUrlMimeType(options.file.mimeType || "text/plain"),
      });

      openBlobInNewTab(blob, pendingWindow);
      return;
    }

    const result = await options.workspace.readFile(options.file.path, {
      mode: "base64",
      maxChars: 16_000_000,
      audit: options.auditContext,
      locationKind: options.fileRef.locationKind,
    });

    if (!result.base64 || result.truncated) {
      throw new Error("File is too large to open in a browser tab.");
    }

    const bytes = base64ToBytes(result.base64);
    const blob = new Blob([toArrayBuffer(bytes)], {
      type: resolveSafeBlobUrlMimeType(options.file.mimeType),
    });

    openBlobInNewTab(blob, pendingWindow);
  } catch (error) {
    closeWindowSafely(pendingWindow);
    throw error;
  }
}

export function createFilesDialogDetailActions(options: CreateFilesDialogDetailActionsOptions): HTMLDivElement {
  const actions = document.createElement("div");
  actions.className = "pi-files-detail-actions";

  const isBuiltIn = isFilesDialogBuiltInDoc(options.file);

  if (isBuiltIn && options.file.kind === "text") {
    // Built-in docs: use clipboard + data-URI download instead of
    // window.open / blob URLs which silently fail in the Office WebView.
    // This add-in always runs inside the Office WebView (loaded via
    // manifest.xml into Excel's sidebar), so this path covers all
    // production usage. Dev-server testing in a browser is unaffected
    // because built-in docs are only available via the workspace.
    const copyButton = document.createElement("button");
    copyButton.type = "button";
    copyButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact";
    copyButton.textContent = t("files-dialog-actions.copyContent");
    copyButton.addEventListener("click", () => {
      void (async () => {
        const result = await options.workspace.readFile(options.file.path, {
          mode: "text",
          maxChars: 16_000_000,
          audit: options.auditContext,
          locationKind: options.fileRef.locationKind,
        });
        if (result.text === undefined) throw new Error("Could not read file.");
        await copyTextToClipboard(result.text, options.file.name);
      })().catch((error: DynamicValue) => {
        showToast(t("files-dialog-actions.toast.copyFailed", { error: getErrorMessage(error) }));
      });
    });

    const downloadButton = document.createElement("button");
    downloadButton.type = "button";
    downloadButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact";
    downloadButton.textContent = t("files-dialog-actions.download");
    downloadButton.addEventListener("click", () => {
      void (async () => {
        const result = await options.workspace.readFile(options.file.path, {
          mode: "text",
          maxChars: 16_000_000,
          audit: options.auditContext,
          locationKind: options.fileRef.locationKind,
        });
        if (result.text === undefined) throw new Error("Could not read file.");
        downloadViaDataUri(
          result.text,
          options.file.name,
          resolveSafeBlobUrlMimeType(options.file.mimeType || "text/plain"),
        );
      })().catch((error: DynamicValue) => {
        showToast(t("files-dialog-actions.toast.downloadFailed", { error: getErrorMessage(error) }));
      });
    });

    actions.append(copyButton, downloadButton);
  } else {
    // Non-built-in files: use the standard blob URL approach.
    const openButton = document.createElement("button");
    openButton.type = "button";
    openButton.className = options.file.kind === "text"
      ? "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact"
      : "pi-overlay-btn pi-overlay-btn--primary pi-overlay-btn--compact";
    openButton.textContent = t("files-dialog-actions.open");
    openButton.addEventListener("click", () => {
      void openFileInBrowser({
        file: options.file,
        fileRef: options.fileRef,
        workspace: options.workspace,
        auditContext: options.auditContext,
      }).catch((error: DynamicValue) => {
        showToast(t("files-dialog-actions.toast.openFailed", { error: getErrorMessage(error) }));
      });
    });

    const downloadButton = document.createElement("button");
    downloadButton.type = "button";
    downloadButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact";
    downloadButton.textContent = t("files-dialog-actions.download");
    downloadButton.addEventListener("click", () => {
      void options.workspace.downloadFile(options.file.path, {
        locationKind: options.fileRef.locationKind,
      }).catch((error: DynamicValue) => {
        showToast(t("files-dialog-actions.toast.downloadFailed", { error: getErrorMessage(error) }));
      });
    });

    actions.append(openButton, downloadButton);
  }

  const isReadOnly = options.file.readOnly || isFilesDialogBuiltInDoc(options.file);
  if (isReadOnly) {
    return actions;
  }

  const renameButton = document.createElement("button");
  renameButton.type = "button";
  renameButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--compact";
  renameButton.textContent = t("files-dialog-actions.rename");
  renameButton.addEventListener("click", () => {
    void (async () => {
      const nextPathInput = await requestTextInputDialog({
        title: t("files-dialog-actions.renameFileTitle"),
        message: `${options.file.path} — leave off the extension to keep it.`,
        initialValue: options.file.path,
        placeholder: "folder/file.ext",
        confirmLabel: t("files-dialog-actions.rename"),
        cancelLabel: t("files-dialog-actions.cancel"),
        restoreFocusOnClose: false,
      });

      if (nextPathInput === null) {
        return;
      }

      const nextPath = resolveRenameDestinationPath(options.file.path, nextPathInput);
      if (nextPath === options.file.path) {
        return;
      }

      await options.workspace.renameFile(options.file.path, nextPath, {
        audit: options.auditContext,
        locationKind: options.fileRef.locationKind,
      });

      showToast(t("files-dialog-actions.toast.renamed", { path: nextPath }));

      await options.onAfterRename(nextPath, options.fileRef.locationKind);
    })().catch((error: DynamicValue) => {
      showToast(t("files-dialog-actions.toast.renameFailed", { error: getErrorMessage(error) }));
    });
  });

  const spacer = document.createElement("div");
  spacer.className = "pi-files-detail-actions__spacer";

  const deleteButton = document.createElement("button");
  deleteButton.type = "button";
  deleteButton.className = "pi-overlay-btn pi-overlay-btn--danger pi-overlay-btn--compact";
  deleteButton.textContent = t("files-dialog-actions.delete");
  deleteButton.addEventListener("click", () => {
    void (async () => {
      const confirmed = await requestConfirmationDialog({
        title: t("files-dialog-actions.deleteFileTitle"),
        message: options.file.path,
        confirmLabel: t("files-dialog-actions.delete"),
        cancelLabel: t("files-dialog-actions.cancel"),
        confirmButtonTone: "danger",
        restoreFocusOnClose: false,
      });

      if (!confirmed) {
        return;
      }

      await options.workspace.deleteFile(options.file.path, {
        audit: options.auditContext,
        locationKind: options.fileRef.locationKind,
      });

      showToast(t("files-dialog-actions.toast.deleted", { name: options.file.name }));

      await options.onAfterDelete();
    })().catch((error: DynamicValue) => {
      showToast(t("files-dialog-actions.toast.deleteFailed", { error: getErrorMessage(error) }));
    });
  });

  actions.append(renameButton, spacer, deleteButton);
  return actions;
}
