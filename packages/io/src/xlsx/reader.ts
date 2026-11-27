import type { SpreadsheetImportResult, SpreadsheetSheet } from '@b32nio/spindle-shared/export/xlsx';
import { ImportError, toImportError } from '../common/import-error';
import { resolveLimits, type ImportLimits } from '../common/safe-zip';
import { abbreviate, CappedWarnings } from './read/capped';
import { XlsxPackage } from './read/package';
import { readWorkbookParts, toImportExtras } from './read';
import { readBudget, readFormats, readSheet } from './read/sheet';
import { findWorkbookPart, isBinaryWorkbook } from './read/workbook';

export interface ImportXlsxOptions {
  /** Overrides of the import limits (`DEFAULT_IMPORT_LIMITS`). */
  limits?: Partial<ImportLimits>;
}

/**
 * Read a workbook with the own package reader (`./read`). The only error this throws is an `ImportError` (code `encrypted`,
 * `legacy-binary`, `not-a-package`, `size-limit` or `corrupt`): the package is opened through the safe zip first, a binary
 * workbook (`.xlsb`) is refused with a message of its own, and whatever the reader itself then fails on is reported as `corrupt`.
 * A part that is damaged or missing is not an error: it is a warning and the rest of the workbook is read.
 *
 * Beyond what the editor uses (`sheets`, `activeSheetIndex`, `warnings`) the result carries the workbook, styles, shared strings
 * and theme, and each sheet its `details`.
 */
export async function importXlsx(bytes: Uint8Array, options: ImportXlsxOptions = {}): Promise<SpreadsheetImportResult> {
  try {
    const pkg = await XlsxPackage.open(bytes, { limits: options.limits });
    const workbookPart = await findWorkbookPart(pkg);
    if (isBinaryWorkbook(pkg, workbookPart)) {
      throw new ImportError('legacy-binary', 'This is a binary Excel workbook (.xlsb), which cannot be opened. In Excel, use Save As and choose Excel Workbook (.xlsx).');
    }
    return await readPackage(pkg, resolveLimits(options.limits));
  } catch (e) {
    throw toImportError(e);
  }
}

async function readPackage(pkg: XlsxPackage, limits: ImportLimits): Promise<SpreadsheetImportResult> {
  // a workbook part that is not XML (or has too many elements) leaves nothing to read (any other damaged part is a warning)
  const parts = await readWorkbookParts(pkg, { strict: true });
  const warnings: string[] = [];
  const sheets: SpreadsheetSheet[] = [];
  // warnings about single sheets (a repeated part, a repeated name) are given for the first few, then counted
  const sheetWarnings = new CappedWarnings((m) => warnings.push(m), 'sheet warnings');
  const names = new Names();
  // what the whole package may still spend, shared by all its sheets, and the parts read so far (a part is read once)
  const budget = readBudget(limits, pkg.loss);
  const readParts = new Set<string>();
  const formats = readFormats(parts.styles, budget, (message) => warnings.push(message));
  // where each workbook sheet (by its place in the workbook) went in `sheets`, for the active tab
  const placed = new Map<number, number>();
  for (const [position, ref] of parts.workbook.sheets.entries()) {
    // only worksheets are read (a chart sheet has no cells), and one whose part is missing was already reported by the package
    if (ref.kind !== 'worksheet' || ref.part === undefined) continue;
    if (sheets.length >= limits.maxSheets) {
      warnings.push(`The workbook has more than ${limits.maxSheets} worksheets (the import limit maxSheets); the others were not read. Raise the limit with the limits option.`);
      pkg.loss.add('sheets', '', parts.workbook.sheets.slice(position).filter((r) => r.kind === 'worksheet' && r.part !== undefined).length);
      break;
    }
    // each part is read once: a second sheet naming a part is how one large part is made to count many times over
    // (the part name is the package's own spelling: the workbook reader has already resolved it, case-insensitively)
    if (readParts.has(ref.part)) {
      sheetWarnings.add(`The sheet "${abbreviate(ref.name)}" names the part ${abbreviate(ref.part)}, which another sheet already uses; it was not read.`);
      continue;
    }
    readParts.add(ref.part);
    const name = names.unique(ref.name, sheetWarnings);
    // a sheet that comes after the cell or scan budget is spent is not kept if it has anything to cut (an empty one is), and a part is
    // still read for that: it is inflated once and stops at its first cell
    const spent = budget.cells <= 0 || budget.scanned <= 0;
    const cuts = budget.cuts;
    const sheet = readSheet(
      await pkg.text(ref.part),
      { ...ref, name },
      {
        styles: parts.styles,
        formats,
        sharedStrings: parts.sharedStrings,
        colors: parts.colors,
        date1904: parts.workbook.date1904,
        maxSheetFormulaChars: limits.maxSheetFormulaChars,
        budget,
        warn: (message) => warnings.push(message),
      }
    );
    if (spent && budget.cuts > cuts) {
      names.forget(name);
      continue;
    }
    placed.set(position, sheets.length);
    sheets.push(sheet);
  }
  sheetWarnings.close();
  // the package's own warnings (damaged parts, unresolvable relationships) come first, in the order they were met
  return {
    sheets,
    activeSheetIndex: activeSheet(parts.workbook.activeTab, placed, sheets.length),
    warnings: [...new Set([...pkg.warnings, ...warnings])],
    ...toImportExtras(parts),
    limitLoss: pkg.loss.items(),
  };
}

/**
 * Sheet names are unique without regard to case (Excel never writes a clash, other producers do): later ones get ` (2)`, ` (3)`, ...
 * and a warning. The next number for each name is kept, so naming many sheets alike costs no more than naming them apart.
 */
class Names {
  private readonly taken = new Set<string>();
  private readonly next = new Map<string, number>();

  unique(name: string, warnings: CappedWarnings): string {
    const key = name.toLowerCase();
    let candidate = name;
    if (this.taken.has(key)) {
      let n = this.next.get(key) ?? 2;
      do candidate = `${name} (${n++})`;
      while (this.taken.has(candidate.toLowerCase()));
      this.next.set(key, n);
      warnings.add(`The workbook has more than one sheet named "${abbreviate(name)}"; one was renamed "${abbreviate(candidate)}".`);
    }
    this.taken.add(candidate.toLowerCase());
    return candidate;
  }

  /** The name of a sheet that was not kept is free again. */
  forget(name: string): void {
    this.taken.delete(name.toLowerCase());
  }
}

/**
 * The editor's active sheet for the workbook's `activeTab` (a place among all sheets, chart sheets and sheets without a part
 * included; none: the editor's default). When that sheet was not read, the next one that was; past the last one, the last. Never outside `sheets`.
 */
function activeSheet(activeTab: number | undefined, placed: ReadonlyMap<number, number>, count: number): number | undefined {
  if (activeTab === undefined) return undefined;
  const tab = Math.max(activeTab, 0);
  let best: number | undefined;
  for (const position of placed.keys()) if (position >= tab && (best === undefined || position < best)) best = position;
  return best === undefined ? Math.max(count - 1, 0) : (placed.get(best) as number);
}
