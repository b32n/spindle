/**
 * Hand bytes (or text) to the browser as a file download. Everything stays
 * local: the data goes into an in-memory Blob + object URL, never over the
 * network.
 */
export function downloadFile(data: Uint8Array | string, fileName: string, mimeType: string): void {
  // Copy into a fresh ArrayBuffer-backed view so the Blob never aliases (or
  // trips over a SharedArrayBuffer-backed) caller-owned buffer.
  const part: BlobPart = typeof data === 'string' ? data : new Uint8Array(data);
  const blob = new Blob([part], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Revoke on the next tick so the browser has started the download.
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

/** Lower-cased extension of `fileName` without the dot ('' if none). */
export function fileExtension(fileName: string): string {
  const dot = fileName.lastIndexOf('.');
  return dot < 0 ? '' : fileName.slice(dot + 1).toLowerCase();
}

/** `fileName` without its final extension. */
export function stripExtension(fileName: string): string {
  const dot = fileName.lastIndexOf('.');
  return dot <= 0 ? fileName : fileName.slice(0, dot);
}

/** Make a title safe to use as a file name (strips path separators etc.). */
export function sanitizeFileName(name: string, fallback = 'Untitled'): string {
  // Control characters are filtered by code point rather than in the regex.
  const cleaned = Array.from(name, (c) => (c.charCodeAt(0) < 32 ? '_' : c))
    .join('')
    .replace(/[\\/:*?"<>|]+/g, '_')
    .trim();
  return cleaned || fallback;
}
