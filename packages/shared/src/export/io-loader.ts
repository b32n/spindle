/**
 * Backward-compatible bridge to `@b32nio/spindle-io`. The XLSX and DOCX
 * readers/writers moved there so this package no longer ships exceljs, docx,
 * jszip and fast-xml-parser; the `export/xlsx` and `export/docx` subpaths here
 * stay as thin async shims that load it on first use.
 *
 * The specifier is a literal so bundlers can see it and split it into a chunk.
 */

const MISSING =
  '@b32nio/spindle-shared/export/%FORMAT% moved to @b32nio/spindle-io. ' +
  'Install it (npm install @b32nio/spindle-io) and import from "@b32nio/spindle-io/%FORMAT%" directly.';

export async function loadIoXlsx() {
  try {
    return await import('@b32nio/spindle-io/xlsx');
  } catch (err) {
    throw new Error(`${MISSING.replace(/%FORMAT%/g, 'xlsx')} (${err instanceof Error ? err.message : String(err)})`);
  }
}

export async function loadIoDocx() {
  try {
    return await import('@b32nio/spindle-io/docx');
  } catch (err) {
    throw new Error(`${MISSING.replace(/%FORMAT%/g, 'docx')} (${err instanceof Error ? err.message : String(err)})`);
  }
}
