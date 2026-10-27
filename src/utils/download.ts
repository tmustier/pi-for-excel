/**
 * Saves a blob as a file from the taskpane.
 *
 * One mechanism for every host: a blob URL on an `<a download>`. Excel for
 * Mac's WKWebView shows its Save sheet for it, even without user activation,
 * but only if nothing in the same task first called `window.open()` on a blob
 * URL; a failed open silently cancels the download. Data URLs and
 * `window.open()` save nothing there, and on Excel on the web `window.open()`
 * leaves a stray tab. The page cannot observe whether the user saved the
 * file, so callers should say the download started, not that it finished.
 */
export function saveBlobAsFile(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.rel = "noopener";
  anchor.hidden = true;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();

  // Excel for Mac reads the blob when its Save sheet opens, seconds later.
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}
