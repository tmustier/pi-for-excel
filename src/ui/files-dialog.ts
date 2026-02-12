/**
 * Files workspace dialog.
 */

import { isExperimentalFeatureEnabled, setExperimentalFeatureEnabled } from "../experiments/flags.js";
import { formatBytes } from "../files/mime.js";
import { FILES_WORKSPACE_CHANGED_EVENT, type WorkspaceFileEntry } from "../files/types.js";
import { getFilesWorkspace } from "../files/workspace.js";
import { getErrorMessage } from "../utils/errors.js";
import { showToast } from "./toast.js";

const OVERLAY_ID = "pi-files-workspace-overlay";

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

  const editor = document.createElement("div");
  editor.className = "pi-files-dialog__editor";
  editor.hidden = true;

  const editorHeader = document.createElement("div");
  editorHeader.className = "pi-files-dialog__editor-header";

  const editorTitle = document.createElement("div");
  editorTitle.className = "pi-files-dialog__editor-title";

  const editorActions = document.createElement("div");
  editorActions.className = "pi-files-dialog__editor-actions";

  const saveButton = makeButton("Save", "pi-files-dialog__btn pi-files-dialog__btn--primary");
  const closeEditorButton = makeButton("Close", "pi-files-dialog__btn");

  const editorNote = document.createElement("div");
  editorNote.className = "pi-files-dialog__editor-note";

  const editorTextarea = document.createElement("textarea");
  editorTextarea.className = "pi-files-dialog__textarea";
  editorTextarea.spellcheck = false;

  editorActions.append(saveButton, closeEditorButton);
  editorHeader.append(editorTitle, editorActions);
  editor.append(editorHeader, editorNote, editorTextarea);

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
    editor,
    footer,
  );

  overlay.appendChild(card);

  let activeEditorPath: string | null = null;
  let editorTruncated = false;

  const closeOverlay = () => {
    cleanup();
    overlay.remove();
  };

  const setStatus = (message: string) => {
    statusLine.textContent = message;
  };

  const clearEditor = () => {
    activeEditorPath = null;
    editorTruncated = false;
    editor.hidden = true;
    editorTitle.textContent = "";
    editorNote.textContent = "";
    editorTextarea.value = "";
    editorTextarea.disabled = false;
    saveButton.disabled = false;
  };

  const openEditor = async (entry: WorkspaceFileEntry) => {
    try {
      const result = await workspace.readFile(entry.path, {
        mode: "text",
        maxChars: 1_000_000,
      });

      editor.hidden = false;
      activeEditorPath = entry.path;
      editorTruncated = result.truncated === true;
      editorTitle.textContent = entry.path;
      editorTextarea.value = result.text ?? "";
      editorTextarea.disabled = editorTruncated;
      saveButton.disabled = editorTruncated;
      editorNote.textContent = editorTruncated
        ? "This file is too large to edit inline safely (preview truncated to 1,000,000 chars)."
        : "Editable text file.";
    } catch (error: unknown) {
      editor.hidden = false;
      activeEditorPath = null;
      editorTruncated = false;
      editorTitle.textContent = entry.path;
      editorTextarea.value = "";
      editorTextarea.disabled = true;
      saveButton.disabled = true;
      editorNote.textContent = `Preview unavailable: ${getErrorMessage(error)}`;
    }
  };

  const renderList = async () => {
    const [backend, files] = await Promise.all([
      workspace.getBackendStatus(),
      workspace.listFiles(),
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
      return;
    }

    for (const file of files) {
      const row = document.createElement("div");
      row.className = "pi-files-dialog__row";

      const info = document.createElement("div");
      info.className = "pi-files-dialog__info";

      const name = document.createElement("div");
      name.className = "pi-files-dialog__name";
      name.textContent = file.path;

      const meta = document.createElement("div");
      meta.className = "pi-files-dialog__meta";
      meta.textContent = `${formatBytes(file.size)} · ${file.kind} · ${formatRelativeDate(file.modifiedAt)}`;

      info.append(name, meta);

      const actions = document.createElement("div");
      actions.className = "pi-files-dialog__actions";

      const editButton = makeButton("Open", "pi-files-dialog__row-btn");
      editButton.addEventListener("click", () => {
        void openEditor(file);
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

        void workspace.renameFile(file.path, nextName).catch((error: unknown) => {
          showToast(`Rename failed: ${getErrorMessage(error)}`);
        });
      });

      const deleteButton = makeButton("Delete", "pi-files-dialog__row-btn pi-files-dialog__row-btn--danger");
      deleteButton.addEventListener("click", () => {
        const ok = window.confirm(`Delete '${file.path}'?`);
        if (!ok) return;

        void workspace.deleteFile(file.path).catch((error: unknown) => {
          showToast(`Delete failed: ${getErrorMessage(error)}`);
        });
      });

      actions.append(editButton, downloadButton, renameButton, deleteButton);
      row.append(info, actions);
      list.appendChild(row);
    }
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

    void workspace.importFiles(selectedFiles)
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

    void workspace.writeTextFile(path, "")
      .then(() => {
        showToast(`Created ${path}.`);
      })
      .catch((error: unknown) => {
        showToast(`Create failed: ${getErrorMessage(error)}`);
      });
  });

  nativeButton.addEventListener("click", () => {
    void workspace.connectNativeDirectory()
      .then(() => {
        showToast("Connected local folder.");
      })
      .catch((error: unknown) => {
        showToast(`Could not connect folder: ${getErrorMessage(error)}`);
      });
  });

  disconnectNativeButton.addEventListener("click", () => {
    void workspace.disconnectNativeDirectory()
      .then(() => {
        showToast("Switched to sandboxed workspace.");
      })
      .catch((error: unknown) => {
        showToast(`Could not switch workspace: ${getErrorMessage(error)}`);
      });
  });

  saveButton.addEventListener("click", () => {
    if (!activeEditorPath || editorTruncated) return;

    const path = activeEditorPath;
    const nextContent = editorTextarea.value;
    void workspace.writeTextFile(path, nextContent)
      .then(() => {
        showToast(`Saved ${path}.`);
      })
      .catch((error: unknown) => {
        showToast(`Save failed: ${getErrorMessage(error)}`);
      });
  });

  closeEditorButton.addEventListener("click", () => {
    clearEditor();
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
