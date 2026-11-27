import { readdirSync, readFileSync } from 'node:fs';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { resolve } from 'node:path';
import type { SpreadsheetCell, SpreadsheetCellStyle, SpreadsheetImportResult, SpreadsheetSheet } from '@b32nio/spindle-shared/export/xlsx';
import { ImportError } from '../common/import-error';
import { importXlsx } from './reader';
import { importXlsxExceljs } from './reader-exceljs';
import { isDateFormat } from './read/cell-style';
import { numberFormatCode } from './read/styles';
import { corpus, MAIN_NS, REL_NS, workbookZip } from './read/test-support';
import { exportXlsx } from './writer';

// ---------------------------------------------------------------------------
// The differential test (S2): the own reader (`importXlsx`) against the `exceljs` reader it replaces (`importXlsxExceljs`, kept in
// the tree until S14) over the whole corpus, and over generated workbooks. On everything the old reader produced the output must
// be identical, apart from the differences below. Each is an improvement, named here, and a difference that is not one of them
// fails the test; the corpus run also pins how often each one happens, so a new difference cannot hide among the old ones.
//
//   colourResolved          a theme, indexed or tinted colour (font, fill, border) the old reader dropped is now its #rrggbb
//   sharedFormulaExpanded   a cell of a shared formula group the old reader read as a plain value (it could not find the
//                           master) now has its formula, moved by its distance from the master
//   falsyResult             a formula whose saved result is 0 or FALSE had the value null (the old reader dropped a falsy result)
//   formulaKept             a cell that is a hyperlink and a formula keeps both (the old reader kept the link and the value)
//   linkOnRange             a link on a range (`A1:B3`) reaches every cell of it (the old reader only linked a cell named
//                           alone)
//   builtinFormat           a built-in number format id shows its en-US code (the table of `styles-scan`, as the loss
//                           detector uses it) where the old reader used its own table (`mm-dd-yy` for id 14) or had none
//   dateFloat               a date serial (a cell whose number format is a date format) that went through a JavaScript Date
//                           and back differs by under a millisecond; the new reader returns the number as written
//   escapeKept              a number format code whose backslash escapes are kept where the backslash is what makes the character
//                           a literal (`0\\ \\d`, `0.0\\"`, `0\\.00`): the old reader dropped every escape (its `numfmt-xform.js`
//                           regex), which changed what such a code means; the two differ only by those escapes, and the code
//                           must be one of ESCAPE_IMPROVEMENTS, each with its reason
//   escapedLetterNotDate    in a 1904 workbook, a number whose format has an escaped date letter (`0\\ \\d`) is not moved by
//                           1462 days: the old reader unescaped the code, saw a `d` and took the number for a date
//   newWarning              a warning the old reader did not have, of one of the kinds listed in NEW_WARNINGS (a missing
//                           relationship, a duplicate sheet name, a damaged sheet, a cell that is not a number or a date); any
//                           other new warning is a difference that fails the test
//
// What the old reader could not read at all (it threw) is listed separately: the own reader must read it or refuse it with an
// ImportError. The other way round is a regression: where the old reader read a file the own reader must too, and where both
// throw the error code must be the same. A file for which that does not hold is named in REFUSED_NOW or CODE_DIFFERENCES below,
// with the reason; nothing else is allowed, and an entry that stops being true fails the test.
// ---------------------------------------------------------------------------

interface Entry {
  id: string;
  filename: string;
  generatedAtRuntime?: boolean;
  expectImportError?: boolean;
}
const manifest = JSON.parse(readFileSync(resolve(__dirname, '../../../../fixtures/xlsx/manifest.json'), 'utf8')) as { fixtures: Entry[] };
const committed = manifest.fixtures.filter((f) => !f.generatedAtRuntime);

type Category =
  | 'colourResolved'
  | 'sharedFormulaExpanded'
  | 'falsyResult'
  | 'formulaKept'
  | 'linkOnRange'
  | 'builtinFormat'
  | 'dateFloat'
  | 'escapeKept'
  | 'escapedLetterNotDate'
  | 'newWarning';

/**
 * The codes (as the own reader gives them, backslashes kept) that differ from the old reader's on purpose, each with why. The old
 * reader's code is the same with every `\\x` replaced by `x`. `notDate`: the old code reads as a date (a day, month or hour letter
 * that the escape had made a literal), so in a 1904 workbook the old reader moved its numbers by 1462 days.
 */
const ESCAPE_IMPROVEMENTS: Record<string, { reason: string; notDate?: boolean }> = {
  '0.0\\"': { reason: 'inches: the escaped quote is a literal; unescaped it leaves an unbalanced quote that may make the exported file invalid' },
  '0 \\m': { reason: 'an escaped m is the letter, not a month token', notDate: true },
  '0 \\d': { reason: 'an escaped d is the letter, not a day token', notDate: true },
  '0\\h': { reason: 'an escaped h is the letter, not an hour token', notDate: true },
  '\\*0\\*': { reason: 'escaped asterisks are literals; unescaped they are fill syntax' },
  '\\@': { reason: 'an escaped at sign is a literal, not the text placeholder' },
  '0\\.00': { reason: 'an escaped dot is a literal, not the decimal point' },
  '0\\,0': { reason: 'an escaped comma is a literal, not the thousands separator' },
  '0\\%': { reason: 'an escaped percent sign is a literal, it does not scale by 100' },
  '0\\\\': { reason: 'an escaped backslash is a literal; unescaped it escapes whatever follows' },
  '"C:\\Temp"0': { reason: 'a backslash inside quotes is a backslash, not an escape' },
};

const dropEscapes = (code: string): string => code.replace(/\\(.)/g, '$1');

/** Files the old reader read and the own reader refuses, by id, with why. (Empty rows are not allowed: a regression is not an exception.) */
const REFUSED_NOW: Record<string, string> = {
  'edge-no-workbook':
    'a package whose root relationship names a workbook part that is not there: the old reader returned no sheets (the editor layer then raised "The file contains no worksheets."); the own reader refuses it as corrupt, with a message that says what is missing. A file that cannot be opened either way.',
};

/** Files both readers refuse with different codes, by id: the old code, the new one, and why. */
const CODE_DIFFERENCES: Record<string, { old: string; new: string; reason: string }> = {};

/** The kinds of warning the own reader adds. */
const NEW_WARNINGS: RegExp[] = [
  /^The workbook has more than one sheet named /,
  /^The sheet ".*" has no usable relationship \(/,
  /^The hyperlink at .* names the relationship .*, which the sheet does not have\.$/,
  /^The sheet ".*" is not well-formed XML; it was read as far as the damage/,
  /^[^!]+![A-Z]+\d+: ".*" is not a (number|date); imported as text$/,
];

/** The old reader's name for a built-in number format, and the code the own reader gives it. */
const BUILTIN_RENAMES: Record<string, string> = {
  'mm-dd-yy': 'm/d/yyyy',
  'm/d/yy "h":mm': 'm/d/yyyy h:mm',
  '#,##0 ;(#,##0)': '#,##0_);(#,##0)',
  '#,##0 ;[Red](#,##0)': '#,##0_);[Red](#,##0)',
  '#,##0.00 ;(#,##0.00)': '#,##0.00_);(#,##0.00)',
  '#,##0.00 ;[Red](#,##0.00)': '#,##0.00_);[Red](#,##0.00)',
  mmss: 'mm:ss.0',
  'mmss.0': 'mm:ss.0',
};

/** What the e2e files show of the named improvements: nothing, they are all identical. */
const E2E_IMPROVEMENTS: Record<string, number> = {};

const addressOf = (key: string): string => {
  const [r, c] = key.split(':').map(Number);
  let letters = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  return `${letters}${r + 1}`;
};

/** The old reader's cell with everything the own reader adds on the side removed. */
const plain = (c: SpreadsheetCell | undefined): SpreadsheetCell | undefined => {
  if (c === undefined) return undefined;
  const { styleIndex: _s, sharedString: _h, formulaInfo: _f, ...rest } = c;
  return rest;
};

const COLOUR_KEYS = ['fontColor', 'backgroundColor'] as const;

function compareStyle(o: SpreadsheetCellStyle | undefined, n: SpreadsheetCellStyle | undefined, seen: Category[]): string | undefined {
  const strip = (s: SpreadsheetCellStyle | undefined): Record<string, unknown> => {
    if (s === undefined) return {};
    const { fontColor: _a, backgroundColor: _b, border, ...rest } = s;
    const out: Record<string, unknown> = { ...rest };
    if (border) {
      const sides: Record<string, unknown> = {};
      for (const side of ['top', 'right', 'bottom', 'left'] as const) {
        const b = border[side];
        if (b) sides[side] = { style: b.style };
      }
      out.border = sides;
    }
    return out;
  };
  const so = strip(o);
  const sn = strip(n);
  if (JSON.stringify(so) !== JSON.stringify(sn)) return `style ${JSON.stringify(so)} vs ${JSON.stringify(sn)}`;
  // colours: what the old reader had, the new one has too, and it may have more
  let added = false;
  for (const key of COLOUR_KEYS) {
    if (o?.[key] !== undefined && o[key] !== n?.[key]) return `${key} ${o[key]} vs ${n?.[key]}`;
    if (o?.[key] === undefined && n?.[key] !== undefined) added = true;
  }
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const oc = o?.border?.[side]?.color;
    const nc = n?.border?.[side]?.color;
    if (oc !== undefined && oc !== nc) return `border ${side} colour ${oc} vs ${nc}`;
    if (oc === undefined && nc !== undefined) added = true;
  }
  if (added) seen.push('colourResolved');
  return undefined;
}

/** A day, in serial units, is 86,400,000 ms: a round trip through a JavaScript Date is good to a millisecond. */
const ONE_MS = 1 / 86_400_000;

function compareCell(
  key: string,
  sheet: SpreadsheetSheet,
  o: SpreadsheetCell,
  n: SpreadsheetCell,
  newStyleIndex: number | undefined,
  newStyles: SpreadsheetImportResult['styles'],
  oldWarnings: string[],
  seen: Category[]
): string | undefined {
  const at = `${sheet.name}!${addressOf(key)}`;
  if (o.value !== n.value) {
    if (typeof o.value === 'number' && typeof n.value === 'number' && isDateFormat(n.numberFormat?.code) && Math.abs(o.value - n.value) <= ONE_MS) seen.push('dateFloat');
    else if (
      typeof o.value === 'number' &&
      typeof n.value === 'number' &&
      o.value - n.value === 1462 &&
      n.numberFormat !== undefined &&
      ESCAPE_IMPROVEMENTS[n.numberFormat.code]?.notDate === true &&
      isDateFormat(dropEscapes(n.numberFormat.code))
    )
      seen.push('escapedLetterNotDate');
    else if (o.value === null && (n.value === 0 || n.value === false || n.value === '') && o.formula === n.formula) seen.push('falsyResult');
    else return `${at}: value ${JSON.stringify(o.value)} vs ${JSON.stringify(n.value)}`;
  }
  if (o.formula !== n.formula) {
    if (o.formula === undefined && n.formula !== undefined && oldWarnings.some((w) => w.startsWith(`${at}: shared formula`))) seen.push('sharedFormulaExpanded');
    else if (o.formula === undefined && n.formula !== undefined && o.hyperlink !== undefined) seen.push('formulaKept');
    else return `${at}: formula ${o.formula} vs ${n.formula}`;
  }
  if (o.hyperlink !== n.hyperlink) {
    const ranged = (sheet.details?.hyperlinks ?? []).some((h) => h.ref.includes(':') && h.target === n.hyperlink);
    if (o.hyperlink === undefined && ranged) seen.push('linkOnRange');
    else return `${at}: hyperlink ${o.hyperlink} vs ${n.hyperlink}`;
  }
  const oc = o.numberFormat?.code;
  const nc = n.numberFormat?.code;
  if (oc !== nc) {
    // a built-in id: the new code is the table's code for the id of the cell's format, and the old one was missing or its own name for it
    const id = newStyleIndex === undefined ? undefined : newStyles?.cellXfs[newStyleIndex]?.numFmtId;
    const builtin = id === undefined || newStyles === undefined || newStyles.numFmts.some((f) => f.id === id) ? undefined : numberFormatCode(newStyles, id);
    if (nc !== undefined && nc === builtin && (oc === undefined || BUILTIN_RENAMES[oc] === nc)) seen.push('builtinFormat');
    else if (nc !== undefined && oc !== undefined && ESCAPE_IMPROVEMENTS[nc] !== undefined && dropEscapes(nc) === oc) seen.push('escapeKept');
    else return `${at}: number format ${oc} vs ${nc}`;
  }
  const style = compareStyle(o.style, n.style, seen);
  return style === undefined ? undefined : `${at}: ${style}`;
}

/** Differences between the two results: those that are not one of the named improvements, and the named ones that were seen. */
function compare(o: SpreadsheetImportResult, n: SpreadsheetImportResult): { unexplained: string[]; seen: Category[] } {
  const unexplained: string[] = [];
  const seen: Category[] = [];
  if (o.activeSheetIndex !== n.activeSheetIndex) unexplained.push(`active sheet ${o.activeSheetIndex} vs ${n.activeSheetIndex}`);
  if (o.sheets.length !== n.sheets.length) return { unexplained: [`${o.sheets.length} sheets vs ${n.sheets.length}`], seen };
  o.sheets.forEach((os, i) => {
    const ns = n.sheets[i];
    if (os.name !== ns.name) unexplained.push(`sheet ${i} name ${os.name} vs ${ns.name}`);
    const keys = new Set([...os.cells.keys(), ...ns.cells.keys()]);
    for (const key of keys) {
      const oc = plain(os.cells.get(key));
      const nc = plain(ns.cells.get(key));
      const newStyleIndex = ns.cells.get(key)?.styleIndex;
      if (oc === undefined || nc === undefined) unexplained.push(`${os.name}!${addressOf(key)}: ${oc === undefined ? 'only in the new reader' : 'only in the old reader'}`);
      else {
        const problem = compareCell(key, ns, oc, nc, newStyleIndex, n.styles, o.warnings, seen);
        if (problem) unexplained.push(problem);
      }
    }
    const asJson = (v: unknown) => JSON.stringify(v, (_k, x) => (x instanceof Map ? [...x] : x instanceof Set ? [...x] : x));
    for (const field of ['columnWidths', 'rowHeights', 'hiddenCols', 'hiddenRows', 'frozenRows', 'frozenCols', 'mergedRanges'] as const) {
      if (asJson(os[field] ?? null) !== asJson(ns[field] ?? null)) unexplained.push(`${os.name}.${field} ${asJson(os[field])} vs ${asJson(ns[field])}`);
    }
  });
  // warnings: those of the old reader that are still true remain; the shared-formula ones go with the formulas they were about
  for (const w of o.warnings) {
    if (!n.warnings.includes(w) && !/shared formula without a resolvable expression/.test(w)) unexplained.push(`old warning missing: ${w}`);
  }
  for (const w of n.warnings.filter((x) => !o.warnings.includes(x))) {
    if (NEW_WARNINGS.some((re) => re.test(w))) seen.push('newWarning');
    else unexplained.push(`new warning of no listed kind: ${w}`);
  }
  return { unexplained, seen };
}

const tally = (seen: Category[]): Record<string, number> => {
  const out: Record<string, number> = {};
  for (const c of seen) out[c] = (out[c] ?? 0) + 1;
  return out;
};

describe('differential: own reader against the exceljs reader, over the corpus', () => {
  const both = async (bytes: Uint8Array) => {
    const o = await importXlsxExceljs(bytes).then((r) => ({ r }), (e: unknown) => ({ e }));
    const n = await importXlsx(bytes).then((r) => ({ r }), (e: unknown) => ({ e }));
    return { o, n };
  };

  it('has the corpus it expects', () => {
    expect(committed.length).toBeGreaterThanOrEqual(40);
  });

  it.each(committed.map((f) => [f.id, f.filename]))('%s: identical output where the old reader had output, apart from the named improvements', async (id, filename) => {
    const { o, n } = await both(corpus(filename));
    // the own reader never throws anything but an ImportError
    if ('e' in n) expect(n.e).toBeInstanceOf(ImportError);
    if ('e' in o) {
      // the old reader threw: the own reader reads the file, or refuses it with the same code (unless the difference is listed)
      if ('e' in n) {
        const listed = CODE_DIFFERENCES[id];
        const codes = { old: (o.e as ImportError).code, new: (n.e as ImportError).code };
        if (listed === undefined) expect(codes.new).toBe(codes.old);
        else expect(codes).toEqual({ old: listed.old, new: listed.new });
      } else expect(CODE_DIFFERENCES[id]).toBeUndefined();
      return;
    }
    if ('e' in n) {
      // the old reader read the file: the own reader must too, unless the file is listed with the reason
      expect({ id, refused: REFUSED_NOW[id] === undefined ? (n.e as ImportError).message : REFUSED_NOW[id] }).toEqual({ id, refused: REFUSED_NOW[id] });
      return;
    }
    expect(REFUSED_NOW[id]).toBeUndefined();
    const { unexplained } = compare(o.r, n.r);
    expect(unexplained).toEqual([]);
  });

  it('lists exceptions that are true: each file of REFUSED_NOW and CODE_DIFFERENCES is in the corpus and behaves as listed', async () => {
    const ids = new Map(committed.map((f) => [f.id, f.filename]));
    for (const id of Object.keys(REFUSED_NOW)) {
      expect(ids.has(id)).toBe(true);
      const { o, n } = await both(corpus(ids.get(id) as string));
      expect('r' in o && 'e' in n).toBe(true);
      expect(REFUSED_NOW[id].length).toBeGreaterThan(40);
    }
    for (const [id, listed] of Object.entries(CODE_DIFFERENCES)) {
      expect(ids.has(id)).toBe(true);
      const { o, n } = await both(corpus(ids.get(id) as string));
      expect([(o as { e: ImportError }).e.code, (n as { e: ImportError }).e.code]).toEqual([listed.old, listed.new]);
      expect(listed.reason.length).toBeGreaterThan(40);
    }
  });

  it('compares the error code where both readers throw: the corpus has three such files and they agree', async () => {
    const rows: Array<[string, string, string]> = [];
    for (const f of committed) {
      const { o, n } = await both(corpus(f.filename));
      if ('e' in o && 'e' in n) rows.push([f.id, (o.e as ImportError).code, (n.e as ImportError).code]);
    }
    expect(rows).toEqual([
      ['edge-truncated-zip', 'corrupt', 'corrupt'],
      ['edge-encrypted', 'encrypted', 'encrypted'],
      ['edge-legacy-xls', 'legacy-binary', 'legacy-binary'],
    ]);
  });

  it('pins every named improvement over the corpus', async () => {
    const seen: Category[] = [];
    for (const f of committed) {
      const { o, n } = await both(corpus(f.filename));
      if ('r' in o && 'r' in n) seen.push(...compare(o.r, n.r).seen);
    }
    expect(tally(seen)).toEqual({ builtinFormat: 8, colourResolved: 3, falsyResult: 7, newWarning: 1, sharedFormulaExpanded: 4 });
  });

  it('lists the files the old reader could not read: the own reader reads them or refuses them with an ImportError', async () => {
    const rows: Array<[string, string, string]> = [];
    for (const f of committed) {
      const { o, n } = await both(corpus(f.filename));
      if (!('e' in o)) continue;
      if ('e' in n) {
        expect(n.e).toBeInstanceOf(ImportError);
        rows.push([f.id, 'old throws', `new refuses (${(n.e as ImportError).code})`]);
      } else rows.push([f.id, 'old throws', `new reads ${n.r.sheets.length} sheet(s), ${n.r.sheets.reduce((a, s) => a + s.cells.size, 0)} cell(s)`]);
    }
    expect(rows).toEqual([
      ['openpyxl-image', 'old throws', 'new reads 1 sheet(s), 25 cell(s)'],
      ['openpyxl-chart', 'old throws', 'new reads 1 sheet(s), 25 cell(s)'],
      ['openpyxl-comment', 'old throws', 'new reads 1 sheet(s), 25 cell(s)'],
      ['edge-truncated-zip', 'old throws', 'new refuses (corrupt)'],
      ['edge-invalid-xml', 'old throws', 'new reads 1 sheet(s), 0 cell(s)'],
      ['edge-encrypted', 'old throws', 'new refuses (encrypted)'],
      ['edge-legacy-xls', 'old throws', 'new refuses (legacy-binary)'],
      ['edge-duplicate-sheet-names', 'old throws', 'new reads 3 sheet(s), 3 cell(s)'],
      ['edge-index-out-of-range', 'old throws', 'new reads 1 sheet(s), 2 cell(s)'],
    ]);
  });

  it('shows the improvements the corpus has no example of: a link on a formula cell, a link on a range, a cached empty string', async () => {
    const bytes = await workbookZip(
      '<sheets><sheet name="L" sheetId="1" r:id="rId1"/></sheets>',
      [['rId1', 'worksheet', 'worksheets/sheet1.xml']],
      {
        'xl/worksheets/sheet1.xml':
          `<worksheet ${MAIN_NS} ${REL_NS}><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><f>A1+1</f><v>2</v></c><c r="D1" t="str"><f>"x"&amp;""</f><v></v></c></row></sheetData>` +
          '<hyperlinks><hyperlink ref="A1:B1" r:id="rId1"/><hyperlink ref="C1" r:id="rId1"/></hyperlinks></worksheet>',
        'xl/worksheets/_rels/sheet1.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>',
      }
    );
    const o = await importXlsxExceljs(bytes);
    const n = await importXlsx(bytes);
    const { unexplained, seen } = compare(o, n);
    expect(unexplained).toEqual([]);
    expect(tally(seen)).toEqual({ formulaKept: 1, linkOnRange: 2, falsyResult: 1 });
  });

  it('reads a frozenSplit pane (ECMA-376 18.3.1.66: frozen, split before it was frozen) as frozen, as the old reader does: no difference to name', async () => {
    const bytes = await workbookZip(
      '<sheets><sheet name="F" sheetId="1" r:id="rId1"/></sheets>',
      [['rId1', 'worksheet', 'worksheets/sheet1.xml']],
      { 'xl/worksheets/sheet1.xml': `<worksheet ${MAIN_NS} ${REL_NS}><sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozenSplit"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>` }
    );
    const o = await importXlsxExceljs(bytes);
    const n = await importXlsx(bytes);
    expect([o.sheets[0].frozenRows, o.sheets[0].frozenCols]).toEqual([3, 2]);
    expect([n.sheets[0].frozenRows, n.sheets[0].frozenCols]).toEqual([3, 2]);
    expect(compare(o, n)).toEqual({ unexplained: [], seen: [] });
  });

  it('imports the three openpyxl files that made the old reader throw (probe 1)', async () => {
    for (const filename of ['openpyxl-image.xlsx', 'openpyxl-chart.xlsx', 'openpyxl-comment.xlsx']) {
      const { o, n } = await both(corpus(filename));
      expect('e' in o).toBe(true);
      expect('r' in n).toBe(true);
      if ('r' in n) expect(n.r.sheets[0].cells.get('0:0')).toMatchObject({ value: 'Region' });
    }
  });
});

// ---------------------------------------------------------------------------
// The files the end-to-end specs open (e2e/fixtures, and the workbooks e2e/sheets build in the test itself). The corpus alone let a
// regression through: LibreOffice escapes every literal in a number format (`\$#,##0.00`), no corpus file does, and the old reader
// dropped the backslashes the editor's format mapping expects. These run the same comparison, with no named improvement beyond
// the corpus ones, and pin the improvements that show up.
// ---------------------------------------------------------------------------

const E2E_FIXTURES = resolve(__dirname, '../../../../e2e/fixtures');
const e2eFiles = readdirSync(E2E_FIXTURES).filter((f) => /\.(xlsx|xlsm)$/i.test(f)).sort();
const e2eBytes = (filename: string): Uint8Array => new Uint8Array(readFileSync(resolve(E2E_FIXTURES, filename)));

/** The cases of e2e/sheets/number-formats.spec.ts, as that spec writes them (exceljs writer). */
const E2E_FORMAT_CASES: Array<[number, string]> = [
  [1234.5, '"$"#,##0.00'],
  [46024, 'dddd, mmmm d, yyyy'],
  [46024.75, 'h:mm AM/PM'],
  [1.5, '[h]:mm:ss'],
  [1500000, '0.0,,"M"'],
  [150, '[>100]"big";"small"'],
];

async function e2eWorkbooks(): Promise<Array<[string, Uint8Array]>> {
  const formats = new ExcelJS.Workbook();
  const sheet = formats.addWorksheet('Formats');
  E2E_FORMAT_CASES.forEach(([value, code], row) => {
    const cell = sheet.getCell(row + 1, 1);
    cell.value = value;
    cell.numFmt = code;
  });
  // e2e/sheets/import-loss.spec.ts: the frozen-header fixture with a conditional format and a chart bolted on
  const zip = await JSZip.loadAsync(e2eBytes('exceljs-frozen.xlsx'));
  const path = Object.keys(zip.files).find((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)) as string;
  const xml = await (zip.file(path) as JSZip.JSZipObject).async('string');
  zip.file(path, xml.replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>'));
  zip.file('xl/charts/chart1.xml', '<c:chartSpace/>');
  return [
    ['number-formats.spec: formats.xlsx', new Uint8Array(await formats.xlsx.writeBuffer())],
    ['import-loss.spec: features.xlsx', await zip.generateAsync({ type: 'uint8array' })],
  ];
}

describe('differential: the files the end-to-end specs open', () => {
  it('finds the e2e fixtures', () => {
    expect(e2eFiles).toEqual(['exceljs-frozen.xlsx', 'styled-multisheet.xlsx']);
  });

  const check = async (name: string, bytes: Uint8Array) => {
    const o = await importXlsxExceljs(bytes);
    const n = await importXlsx(bytes);
    const { unexplained, seen } = compare(o, n);
    expect({ name, unexplained }).toEqual({ name, unexplained: [] });
    return seen;
  };

  it.each(e2eFiles)('%s: identical output apart from the named improvements', async (filename) => {
    await check(filename, e2eBytes(filename));
  });

  it('the workbooks the e2e specs build: identical output apart from the named improvements', async () => {
    for (const [name, bytes] of await e2eWorkbooks()) await check(name, bytes);
  });

  it('pins the improvements the e2e files show', async () => {
    const seen: Category[] = [];
    for (const f of e2eFiles) seen.push(...(await check(f, e2eBytes(f))));
    for (const [name, bytes] of await e2eWorkbooks()) seen.push(...(await check(name, bytes)));
    expect(tally(seen)).toEqual(E2E_IMPROVEMENTS);
  });

  it('S-FIO-01: the LibreOffice file keeps the codes LibreOffice escapes, as the old reader unescaped them', async () => {
    const n = await importXlsx(e2eBytes('styled-multisheet.xlsx'));
    const sales = n.sheets.find((s) => s.name === 'Sales') as SpreadsheetSheet;
    expect(sales.cells.get('1:2')?.numberFormat).toEqual({ code: '$#,##0.00' });
    expect(sales.cells.get('1:5')?.numberFormat).toEqual({ code: 'yyyy-mm-dd' });
    expect(sales.cells.get('1:6')?.numberFormat).toEqual({ code: '#,##0.0" kg"' });
    // the workbook's styles keep the code as the file writes it
    expect(n.styles?.numFmts.find((f) => f.id === 165)?.code).toBe('\\$#,##0.00');
  });
});

describe('differential: the escapes the own reader keeps (escapeKept)', () => {
  const NS = `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"`;
  const attr = (code: string): string => code.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  /** A one-cell workbook with the number format `code` written as given and the value 7, in the 1900 or 1904 date system. */
  const book = (code: string, date1904: boolean) =>
    workbookZip(
      '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>',
      [['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'styles', 'styles.xml']],
      {
        'xl/workbook.xml': `<workbook ${NS} ${REL_NS}><workbookPr${date1904 ? ' date1904="1"' : ''}/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        'xl/styles.xml': `<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="170" formatCode="${attr(code)}"/></numFmts><cellXfs><xf/><xf numFmtId="170"/></cellXfs></styleSheet>`,
        'xl/worksheets/sheet1.xml': `<worksheet ${NS}><sheetData><row r="1"><c r="A1" s="1"><v>7</v></c></row></sheetData></worksheet>`,
      }
    );

  it.each(Object.entries(ESCAPE_IMPROVEMENTS))('%s: %s', async (code, { reason, notDate }) => {
    expect(reason.length).toBeGreaterThan(10);
    for (const date1904 of [false, true]) {
      const bytes = await book(code, date1904);
      const o = await importXlsxExceljs(bytes);
      const n = await importXlsx(bytes);
      const { unexplained, seen } = compare(o, n);
      expect(unexplained).toEqual([]);
      // the old code is the new one without its escapes, and the new one is kept as written (no escape of a plain character in these)
      expect(o.sheets[0].cells.get('0:0')?.numberFormat?.code).toBe(dropEscapes(code));
      expect(n.sheets[0].cells.get('0:0')?.numberFormat?.code).toBe(code);
      expect(seen).toEqual(date1904 && notDate ? ['escapedLetterNotDate', 'escapeKept'] : ['escapeKept']);
      expect(n.sheets[0].cells.get('0:0')?.value).toBe(7);
      if (date1904 && notDate) expect(o.sheets[0].cells.get('0:0')?.value).toBe(7 + 1462);
    }
  });

  it('is no wider than the table: a code that is not listed, or that differs by more than its escapes, or a 1462 difference without a date letter, is a difference', () => {
    const one = (value: number, code?: string): SpreadsheetImportResult => ({ sheets: [{ name: 'S', cells: new Map([['0:0', { value, ...(code === undefined ? {} : { numberFormat: { code } }) }]]) }], warnings: [] });
    expect(compare(one(1, '0.0"'), one(1, '0.0\\"')).seen).toEqual(['escapeKept']);
    expect(compare(one(1, '0 d'), one(1, '0 \\d')).seen).toEqual(['escapeKept']);
    expect(compare(one(1463, '0 d'), one(1, '0 \\d')).seen).toEqual(['escapedLetterNotDate', 'escapeKept']);
    // not listed
    expect(compare(one(1, '0 q'), one(1, '0 \\q')).unexplained).toHaveLength(1);
    // listed, but the old code is not the new one without its escapes
    expect(compare(one(1, '0.1'), one(1, '0.0\\"')).unexplained).toHaveLength(1);
    // no old code at all
    expect(compare(one(1), one(1, '0.0\\"')).unexplained).toHaveLength(1);
    // a 1462 difference where the code is not listed as a non-date, or has no date letter once unescaped
    expect(compare(one(1463, '0 q'), one(1, '0 \\q')).unexplained.length).toBeGreaterThan(0);
    expect(compare(one(1463, '0.0"'), one(1, '0.0\\"')).unexplained.length).toBeGreaterThan(0);
    expect(compare(one(1463), one(1)).unexplained).toHaveLength(1);
    expect(compare(one(1464, '0 d'), one(1, '0 \\d')).unexplained.length).toBeGreaterThan(0);
    expect(compare(one(1463, '0 d'), { sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1 }]]) }], warnings: [] }).unexplained.length).toBeGreaterThan(0);
  });
});

describe('differential: the named improvements are no wider than their names', () => {
  const xfs = (numFmtId: number) => ({ numFmtId } as never);
  const styles = { numFmts: [{ id: 164, code: '0.0" kg"' }], cellXfs: [xfs(0), xfs(14), xfs(164), xfs(4)] } as unknown as NonNullable<SpreadsheetImportResult['styles']>;
  const result = (cell: SpreadsheetCell, extra: Partial<SpreadsheetSheet> = {}, warnings: string[] = [], withStyles = true): SpreadsheetImportResult => ({
    sheets: [{ name: 'S', cells: new Map([['0:0', cell]]), ...extra }],
    warnings,
    ...(withStyles ? { styles } : {}),
  });
  const diff = (o: SpreadsheetCell, n: SpreadsheetCell, extra: Partial<SpreadsheetSheet> = {}, ow: string[] = [], nw: string[] = []) => compare(result(o, {}, ow, false), result(n, extra, nw));

  it('dateFloat: only a date-formatted cell, and only by under a millisecond', () => {
    const date = { numberFormat: { code: 'yyyy-mm-dd' } };
    expect(diff({ value: 45000.123456789, ...date }, { value: 45000.123456789 + 1e-9, ...date })).toEqual({ unexplained: [], seen: ['dateFloat'] });
    expect(diff({ value: 45000.5 }, { value: 45000.5 + 1e-9 }).unexplained).toHaveLength(1);
    expect(diff({ value: 45000.5, numberFormat: { code: '0.00' } }, { value: 45000.5 + 1e-9, numberFormat: { code: '0.00' } }).unexplained).toHaveLength(1);
    expect(diff({ value: 45000.5, ...date }, { value: 45000.5 + 1e-6, ...date }).unexplained).toHaveLength(1);
    expect(diff({ value: 1, ...date }, { value: 2, ...date }).unexplained).toHaveLength(1);
  });

  it('builtinFormat: only the table code of the built-in id of the cell, where the old code was missing or its own name', () => {
    // id 14 is m/d/yyyy; the old reader said mm-dd-yy, or nothing
    expect(diff({ value: 1, numberFormat: { code: 'mm-dd-yy' } }, { value: 1, styleIndex: 1, numberFormat: { code: 'm/d/yyyy' } })).toEqual({ unexplained: [], seen: ['builtinFormat'] });
    expect(diff({ value: 1 }, { value: 1, styleIndex: 1, numberFormat: { code: 'm/d/yyyy' } }).seen).toEqual(['builtinFormat']);
    // any other new code is a difference: a wrong one for the id, an old code that is not its rename, no format index, a custom format, an unknown index
    expect(diff({ value: 1 }, { value: 1, styleIndex: 1, numberFormat: { code: 'dd/mm/yyyy' } }).unexplained).toHaveLength(1);
    expect(diff({ value: 1, numberFormat: { code: '0.00' } }, { value: 1, styleIndex: 1, numberFormat: { code: 'm/d/yyyy' } }).unexplained).toHaveLength(1);
    expect(diff({ value: 1 }, { value: 1, numberFormat: { code: 'm/d/yyyy' } }).unexplained).toHaveLength(1);
    expect(diff({ value: 1 }, { value: 1, styleIndex: 2, numberFormat: { code: '0.0" kg"' } }).unexplained).toHaveLength(1);
    expect(diff({ value: 1 }, { value: 1, styleIndex: 9, numberFormat: { code: '0' } }).unexplained).toHaveLength(1);
    expect(compare(result({ value: 1 }, {}, [], false), result({ value: 1, styleIndex: 1, numberFormat: { code: 'm/d/yyyy' } }, {}, [], false)).unexplained).toHaveLength(1);
  });

  it('newWarning: only the listed kinds', () => {
    const kinds = [
      'The workbook has more than one sheet named "A"; one was renamed "A (2)".',
      'The sheet "Dangling" has no usable relationship (rId99).',
      'The hyperlink at A1 names the relationship rId9, which the sheet does not have.',
      'The sheet "S" is not well-formed XML; it was read as far as the damage (3 cells).',
      'S!B2: "12abc" is not a number; imported as text',
      'S!B2: "2020-02-30" is not a date; imported as text',
    ];
    for (const w of kinds) expect(diff({ value: 1 }, { value: 1 }, {}, [], [w])).toEqual({ unexplained: [], seen: ['newWarning'] });
    expect(diff({ value: 1 }, { value: 1 }, {}, [], ['something new']).unexplained).toEqual(['new warning of no listed kind: something new']);
    expect(diff({ value: 1 }, { value: 1 }, {}, ['kept'], ['kept', 'S: formulas left out, their values kept: 1 over 8192 characters.']).unexplained).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// Generated workbooks: a seeded generator writes workbooks through the exceljs writer, so the old reader can read them, and both
// readers must agree. A fixed seed: the same workbooks every run.
// ---------------------------------------------------------------------------

function rng(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = Math.imul(s ^ (s >>> 15), 1 | s);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const FORMATS = ['0', '0.00', '#,##0.00', '0%', 'yyyy-mm-dd', 'm/d/yyyy h:mm', '0.0" kg"', '@', 'h:mm:ss'];
const COLOURS = ['#ff0000', '#00ff00', '#0000ff', '#123456', '#abcdef'];
const BORDERS = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'] as const;
const WORDS = ['alpha', 'beta', 'gamma', 'x & y', '<tag>', 'a "quote"', 'ünï', '  spaced  ', 'line\nbreak'];

function generateSheet(name: string, next: () => number): SpreadsheetSheet {
  const pick = <T>(xs: readonly T[]): T => xs[Math.floor(next() * xs.length)];
  const cells = new Map<string, SpreadsheetCell>();
  const rows = 3 + Math.floor(next() * 12);
  const cols = 2 + Math.floor(next() * 6);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (next() < 0.25) continue;
      const kind = next();
      const cell: SpreadsheetCell = { value: null };
      if (kind < 0.35) cell.value = Math.round(next() * 1e6) / 100 - 2000;
      else if (kind < 0.6) cell.value = pick(WORDS);
      else if (kind < 0.7) cell.value = next() < 0.5;
      else if (kind < 0.8) {
        cell.formula = `SUM(A1:A${r + 1})*${c + 1}`;
        cell.value = next() < 0.2 ? 0 : Math.round(next() * 1000);
      } else if (kind < 0.88) cell.value = pick(WORDS).trim() || 'x';
      else cell.value = Math.floor(next() * 100000);
      if (next() < 0.3) cell.numberFormat = { code: pick(FORMATS) };
      if (next() < 0.5) {
        const style: SpreadsheetCellStyle = {};
        if (next() < 0.3) style.bold = true;
        if (next() < 0.2) style.italic = true;
        if (next() < 0.2) style.underline = true;
        if (next() < 0.1) style.strikethrough = true;
        if (next() < 0.3) style.fontFamily = pick(['Arial', 'Calibri', 'Times New Roman']);
        if (next() < 0.3) style.fontSize = pick([8, 10, 11, 14, 24]);
        if (next() < 0.3) style.fontColor = pick(COLOURS);
        if (next() < 0.3) style.backgroundColor = pick(COLOURS);
        if (next() < 0.3) style.horizontalAlign = pick(['left', 'center', 'right'] as const);
        if (next() < 0.3) style.verticalAlign = pick(['top', 'middle', 'bottom'] as const);
        if (next() < 0.2) style.wrapText = true;
        if (next() < 0.15) style.textRotation = pick([0, 45, 90, -45, -90]);
        if (next() < 0.25) {
          style.border = {};
          for (const side of ['top', 'right', 'bottom', 'left'] as const) if (next() < 0.6) style.border[side] = { style: pick(BORDERS), color: next() < 0.5 ? pick(COLOURS) : undefined };
        }
        if (Object.keys(style).length > 0) cell.style = style;
      }
      if (next() < 0.05 && cell.formula === undefined) cell.hyperlink = `https://example.com/${r}/${c}`;
      cells.set(`${r}:${c}`, cell);
    }
  }
  const sheet: SpreadsheetSheet = { name, cells };
  if (next() < 0.6) {
    sheet.columnWidths = new Map();
    for (let c = 0; c < cols; c++) if (next() < 0.5) sheet.columnWidths.set(c, 40 + Math.floor(next() * 300));
  }
  if (next() < 0.6) {
    sheet.rowHeights = new Map();
    for (let r = 0; r < rows; r++) if (next() < 0.3) sheet.rowHeights.set(r, 16 + Math.floor(next() * 60));
  }
  if (next() < 0.3) sheet.hiddenCols = new Set([Math.floor(next() * cols)]);
  if (next() < 0.3) sheet.hiddenRows = new Set([Math.floor(next() * rows)]);
  if (next() < 0.4) {
    sheet.frozenRows = Math.floor(next() * 3);
    sheet.frozenCols = Math.floor(next() * 3);
  }
  if (next() < 0.4 && rows > 3 && cols > 3) {
    const r = Math.floor(next() * (rows - 2));
    const c = Math.floor(next() * (cols - 2));
    sheet.mergedRanges = [{ startRow: r, startCol: c, endRow: r + 1, endCol: c + 1 }];
  }
  return sheet;
}

describe('differential: own reader against the exceljs reader, over generated workbooks', () => {
  const seen: Category[] = [];

  it.each(Array.from({ length: 40 }, (_, i) => i + 1))('seed %i', async (seed) => {
    const next = rng(seed);
    const sheets = Array.from({ length: 1 + Math.floor(next() * 3) }, (_, i) => generateSheet(`Sheet ${i + 1}`, next));
    const bytes = await exportXlsx({ sheets, activeSheetIndex: Math.floor(next() * sheets.length) });
    const o = await importXlsxExceljs(bytes);
    const n = await importXlsx(bytes);
    const { unexplained, seen: s } = compare(o, n);
    seen.push(...s);
    expect(unexplained).toEqual([]);
  });

  it('only ever needs the named improvements', () => {
    // a falsy cached result and dates whose serial the old reader moved through a Date are what generated workbooks can show
    expect(Object.keys(tally(seen)).sort()).toEqual(['dateFloat', 'falsyResult']);
  });
});
