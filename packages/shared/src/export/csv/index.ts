// Pure CSV read/write: grid <-> RFC4180-ish text. No knowledge of any
// editor's data model — callers resolve their own cells/values down to
// CsvCellValue before calling exportCsv, and interpret importCsv's output
// back into their own model.

export type CsvCellValue = string | number | boolean | null;

export interface CsvExportInput {
  rows: CsvCellValue[][];
}

export interface CsvImportResult {
  rows: (string | number)[][];
}

/**
 * Format a grid of values as CSV text: comma/quote/newline escaping, plus a
 * guard against formula injection (a leading =, +, -, @, tab, or CR makes
 * Excel/Sheets treat a string cell as a formula when the file is opened).
 * Numbers and booleans are never guarded or quoted — they can't carry an
 * injection payload, and quoting a bare number would corrupt it into text on
 * re-import.
 */
export function exportCsv(input: CsvExportInput): string {
  return input.rows.map((row) => row.map(formatCsvField).join(',')).join('\n');
}

function formatCsvField(value: CsvCellValue): string {
  let text = '';
  if (typeof value === 'number' || typeof value === 'boolean') {
    text = value.toString();
  } else if (value != null) {
    text = guardCsvInjection(value);
  }
  if (text.includes(',') || text.includes('"') || text.includes('\n')) {
    text = `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

/** Prefix a leading formula-trigger character with a quote to prevent CSV injection. */
function guardCsvInjection(value: string): string {
  return /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
}

/**
 * Parse CSV text into a grid of values, coercing round-trip-safe numeric
 * strings back to numbers (so exportCsv's unquoted numbers re-import as
 * numbers, per its own contract).
 */
export function importCsv(csv: string): CsvImportResult {
  return { rows: parseCsv(csv).map((row) => row.map(coerceCsvField)) };
}

/**
 * Coerce a CSV field to a number when it round-trips exactly, while
 * preserving as text anything where coercion would lose information —
 * leading zeros ("007"), exponential spellings ("1e21"), or explicit
 * trailing zeros ("5.0").
 */
function coerceCsvField(value: string): string | number {
  const n = Number(value);
  if (value.trim() !== '' && !Number.isNaN(n) && String(n) === value.trim()) {
    return n;
  }
  return value;
}

/**
 * Parse a CSV document into rows of fields. Handles quoted fields containing
 * commas and embedded newlines, escaped quotes (`""`), and CRLF / CR / LF line
 * endings — so it round-trips this module's own `exportCsv` output and does
 * not mangle Windows CSVs (trailing `\r` on the last column) or quoted-newline
 * cells (which a naive line-then-field split would scramble).
 */
function parseCsv(csv: string): string[][] {
  const rows: string[][] = [];
  let current = '';
  let field: string[] = [];
  let inQuotes = false;

  const endField = (): void => {
    field.push(current);
    current = '';
  };
  const endRow = (): void => {
    endField();
    rows.push(field);
    field = [];
  };

  for (let i = 0; i < csv.length; i++) {
    const char = csv[i];
    if (inQuotes) {
      if (char === '"') {
        if (csv[i + 1] === '"') {
          current += '"'; // escaped quote
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += char;
      }
      continue;
    }
    if (char === '"') {
      inQuotes = true;
    } else if (char === ',') {
      endField();
    } else if (char === '\r') {
      // CRLF or bare CR ends a row; swallow the paired LF.
      if (csv[i + 1] === '\n') i++;
      endRow();
    } else if (char === '\n') {
      endRow();
    } else {
      current += char;
    }
  }
  // Flush the final field/row unless the file ended exactly on a row break.
  if (current !== '' || field.length > 0) endRow();
  // Drop a trailing fully-empty row produced by a terminal newline.
  if (rows.length > 0 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') {
    rows.pop();
  }
  return rows;
}
