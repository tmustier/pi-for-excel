# Files Workspace

The Files overlay provides a unified view of all files available to the agent during a session. Files can come from multiple sources depending on the host environment.

## Storage backends

| Backend | Key | When used |
|---------|-----|-----------|
| **OPFS** (Origin Private File System) | `opfs` | Default in all browser-based hosts. Sandboxed, persistent per origin. |
| **Native directory** | `native-directory` | When the user connects a local folder via the File System Access API. |
| **In-memory** | `memory` | Non-browser / test environments. |

The workspace manager selects a backend automatically on startup:

1. If a persisted native directory handle exists and permission is still granted → `native-directory`
2. If OPFS is available → `opfs`
3. Otherwise → `memory`

When a native directory is connected, both the OPFS workspace and the native directory are listed together as dual sources. Files from the connected folder appear in a dedicated section.

## Connected folders

Connected-folder mode lets users link a local directory so the agent can read project files directly. It relies on the **File System Access API** (`showDirectoryPicker`), which is only available in Chromium-based browsers.

### Host compatibility matrix

| Host | Engine | `showDirectoryPicker` | Connect folder button |
|------|--------|----------------------|----------------------|
| Excel Online (Chrome / Edge) | Chromium | ✅ Supported | Visible |
| Excel Online (Firefox) | Gecko | ❌ Not supported | Hidden |
| Excel Online (Safari) | WebKit | ❌ Not supported | Hidden |
| Excel desktop (macOS) | WKWebView | ❌ Not supported | Hidden |
| Excel desktop (Windows) | WebView2 (Chromium) | ⚠️ Varies by version | Auto-detected |

### Behavior in unsupported hosts

When `showDirectoryPicker` is unavailable:

- `backendStatus.nativeSupported` is `false`
- The **Connect folder** button is **hidden** — users never see a broken or confusing control
- The workspace uses OPFS (or in-memory) as the sole storage backend
- All other Files overlay features (upload, download, preview, rename, delete) work normally

### Behavior in supported hosts

When `showDirectoryPicker` is available:

- The **Connect folder** button is visible and enabled
- Clicking it opens the OS directory picker
- Once connected, the button shows **Connected ✓** (disabled)
- Files from the connected directory appear in a separate section labelled with the folder name
- The directory handle is persisted so it can be restored on next session (permission permitting)

## Related

- [PR #356](https://github.com/tmustier/pi-for-excel/pull/356) — Phase 2 dual-source connected-folder sections
- [PR #361](https://github.com/tmustier/pi-for-excel/pull/361) — Hide connect-folder action when unsupported
- [PR #362](https://github.com/tmustier/pi-for-excel/pull/362) — Centralize connect-folder button state
- [Issue #360](https://github.com/tmustier/pi-for-excel/issues/360) — Decision: keep feature, hide when unsupported
