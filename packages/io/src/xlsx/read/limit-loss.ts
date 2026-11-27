import type { ImportLossItem } from '@b32nio/spindle-shared';
import type { ImportLimits } from '../../common/safe-zip';
import { abbreviate } from './capped';

// ---------------------------------------------------------------------------
// What the import limits cut from a file, as loss items. A warning is a line of text for a log; a host shows the import report
// (`ImportLossReport`) to the user, so everything a budget leaves out is also counted here, by kind, and the reader hands the items on
// (`SpreadsheetImportResult.limitLoss`; `importFromXlsx` appends them to the report it returns). One item per kind of cut, all with the
// id `xlsx.package.sizeLimit` (what the detector uses for a package it cannot scan), each naming what was left out, where, and the
// `limits` option that raises the limit.
// ---------------------------------------------------------------------------

export type CutKind =
  | 'cells'
  | 'scan'
  | 'sheets'
  | 'rows'
  | 'ranges'
  | 'columns'
  | 'extLst'
  | 'formulas'
  | 'parts'
  | 'sharedStrings'
  | 'relationships'
  | 'definedNames'
  | 'sheetNames'
  | 'elements'
  | 'numberFormats';

interface Kind {
  label: string;
  /** What happens to what was cut; the first word is `lost` or `approximated`, as in the detector's items. */
  effect: (limits: ImportLimits) => string;
}

const RAISE = 'raise it with the `limits` option of the import';

const KINDS: Record<CutKind, Kind> = {
  cells: { label: 'sheet(s) cut at the cell limit', effect: (l) => `lost: the cells past maxTotalCells (${l.maxTotalCells}) over the workbook were not imported; ${RAISE}` },
  scan: { label: 'sheet(s) cut at the scan limit', effect: (l) => `lost: the rest of the sheet was not read after maxScannedElements (${l.maxScannedElements}) cell, row and column elements over the workbook; ${RAISE}` },
  sheets: { label: 'worksheet(s) not read', effect: (l) => `lost: the sheets past maxSheets (${l.maxSheets}) were not imported (the workbook's entries are read up to four times that); ${RAISE}` },
  rows: { label: 'row(s) with settings left out', effect: (l) => `approximated: the cells are kept, the height, hidden state and style of the rows past maxRows (${l.maxRows}) are not; ${RAISE}` },
  ranges: { label: 'merged range(s) and hyperlink(s) left out', effect: (l) => `lost: past maxRanges (${l.maxRanges}) over the workbook, 1,048,576 merges or 65,530 hyperlinks on a sheet; ${RAISE}` },
  columns: { label: 'column setting(s) left out', effect: (l) => `approximated: the widths and hidden state of the columns past maxColumns (${l.maxColumns}) over the workbook are not kept (16,384 ranges on a sheet); ${RAISE}` },
  extLst: { label: 'extension list(s) not kept', effect: (l) => `lost: the raw extensions (conditional formats, sparklines) past maxExtensionChars (${l.maxExtensionChars}) over the workbook, or over 8 Mi characters on a sheet, are not kept; ${RAISE}` },
  formulas: { label: 'formula(s) not imported', effect: (l) => `approximated: the value is kept, the formula is not (over 8192 characters, or past maxSheetFormulaChars (${l.maxSheetFormulaChars}) or maxFormulaChars (${l.maxFormulaChars}) of expanded shared formulas); ${RAISE}` },
  parts: { label: 'part(s) over the element limit', effect: (l) => `lost: not read, so the reader's defaults are used (more than maxPartElements (${l.maxPartElements}) elements or entries); ${RAISE}` },
  sharedStrings: { label: 'part(s) of the shared strings not read', effect: (l) => `lost: the strings past maxTotalCells (${l.maxTotalCells}) and the formatting past maxPartElements (${l.maxPartElements}) runs are not kept; ${RAISE}` },
  relationships: { label: 'part(s) with relationships not read', effect: (l) => `lost: the relationships past maxRelationships (${l.maxRelationships}) over the workbook were not read, so links and comment parts through them are not found; ${RAISE}` },
  definedNames: { label: 'defined name(s) not kept', effect: (l) => `lost: the names past maxPartElements (${l.maxPartElements}) were not read; ${RAISE}` },
  sheetNames: { label: 'sheet name(s) cut at 255 characters', effect: () => 'approximated: the name is shortened (Excel allows 31 characters)' },
  elements: { label: 'sheet element(s) not read', effect: () => 'lost: a sheet element (properties, dimension, views, format) over 64 KiB was not read, so its defaults are used' },
  numberFormats: { label: 'number format code(s) over 255 characters', effect: () => 'approximated: kept as written, but not read as a date or time (Excel allows 255 characters)' },
};

/** Entries named in an item's `where`. */
const MAX_WHERE = 5;

interface Tally {
  count: number;
  where: string[];
}

/** The cuts of one import, by kind. `add` is called wherever a budget leaves something out; `items` makes the report. */
export class LimitLoss {
  private readonly tallies = new Map<CutKind, Tally>();

  constructor(private readonly limits: ImportLimits) {}

  /** Count `n` (default one) cut things of `kind`, in `where` (a sheet or part name, quoted shortened; the first few are listed; none for a cut that has no place). Returns whether this is the first of its kind. */
  add(kind: CutKind, where = '', n = 1): boolean {
    let tally = this.tallies.get(kind);
    const first = tally === undefined;
    if (tally === undefined) {
      tally = { count: 0, where: [] };
      this.tallies.set(kind, tally);
    }
    tally.count += n;
    const place = abbreviate(where);
    if (where !== '' && tally.where.length < MAX_WHERE && !tally.where.includes(place)) tally.where.push(place);
    return first;
  }

  /** Whether anything of `kind` was cut. */
  has(kind: CutKind): boolean {
    return this.tallies.has(kind);
  }

  items(): ImportLossItem[] {
    const items: ImportLossItem[] = [];
    for (const [kind, { label, effect }] of Object.entries(KINDS) as Array<[CutKind, Kind]>) {
      const tally = this.tallies.get(kind);
      if (tally === undefined) continue;
      const text = effect(this.limits);
      items.push({
        id: 'xlsx.package.sizeLimit',
        label,
        count: tally.count,
        ...(tally.where.length > 0 ? { where: tally.where } : {}),
        effect: text,
        ...(text.startsWith('approximated') ? { state: 'approximated' as const } : {}),
      });
    }
    return items;
  }
}
