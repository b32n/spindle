// Text import (CSV) is synchronous and part of the core bundle, while `ImportError` lives in the lazily loaded
// `@b32nio/spindle-io`, so CSV refusals are errors of the same shape (`name`, `code`, `message`) rather than instances.
// `isImportError` recognises both, so a host can handle every refusal in one place.

const MB = 1024 * 1024;

/** An error that carries an import `code` (`size-limit`, ...): an `ImportError` from `@b32nio/spindle-io`, or the CSV importer's twin. */
export function isImportError(e: unknown): e is Error & { name: 'ImportError'; code: string } {
  return e instanceof Error && e.name === 'ImportError' && typeof (e as { code?: unknown }).code === 'string';
}

/** The refusal of a CSV text or file over `maxChars`, with the limit in the message. */
export function csvTooLargeError(maxChars: number): Error & { name: 'ImportError'; code: 'size-limit' } {
  const limit = maxChars >= MB ? `${Math.round(maxChars / MB)} MB` : `${maxChars} bytes`;
  return Object.assign(new Error(`The file is too large to open safely (the text is over the limit of ${limit}).`), {
    name: 'ImportError' as const,
    code: 'size-limit' as const,
  });
}
