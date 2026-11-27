import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { importXlsxExceljs } from '../reader-exceljs';
import * as api from './index';
import type { SpreadsheetImportResult } from '@b32nio/spindle-shared/export/xlsx';
import { readWorkbookParts, toImportExtras, XlsxPackage } from './index';
import { corpus, workbookZip } from './test-support';

interface Entry {
  id: string;
  filename: string;
  generatedAtRuntime?: boolean;
  expectImportError?: string;
}
const manifest = JSON.parse(readFileSync(resolve(__dirname, '../../../../../fixtures/xlsx/manifest.json'), 'utf8')) as { fixtures: Entry[] };
const committed = manifest.fixtures.filter((f) => !f.generatedAtRuntime);
const readable = committed.filter((f) => !f.expectImportError);

const parts = async (file: string) => readWorkbookParts(await XlsxPackage.open(corpus(file)));

describe('the whole corpus', () => {
  it('has the files this test expects', () => {
    expect(committed.length).toBeGreaterThanOrEqual(40);
    expect(readable.length).toBeGreaterThanOrEqual(30);
  });

  it.each(readable.map((f) => [f.id, f.filename]))('%s: every workbook-level part reads, with the sheets the exceljs reader finds', async (_id, filename) => {
    const bytes = corpus(filename);
    const result = await readWorkbookParts(await XlsxPackage.open(bytes));
    expect(result.workbook.sheets.length).toBeGreaterThan(0);
    // every listed sheet part exists, and every cellXf points inside the tables it indexes
    if (result.warnings.length === 0) for (const sheet of result.workbook.sheets) expect(sheet.part).toBeDefined();
    const { styles } = result;
    for (const xf of styles.cellXfs) {
      expect(xf.fontId).toBeLessThan(Math.max(styles.fonts.length, 1));
      expect(xf.fillId).toBeLessThan(Math.max(styles.fills.length, 1));
      expect(xf.borderId).toBeLessThan(Math.max(styles.borders.length, 1));
    }
    // differential on what the old reader also reports: the sheet names, in order, and the active tab
    let old;
    try {
      old = await importXlsxExceljs(bytes);
    } catch {
      return; // the exceljs reader throws on this file (the openpyxl ones): nothing to compare with
    }
    // (a sheet whose part cannot be found stays listed here with no part; the old reader drops it)
    expect(result.workbook.sheets.filter((s) => s.part !== undefined).map((s) => s.name)).toEqual(old.sheets.map((s) => s.name));
    if (old.activeSheetIndex !== undefined) expect(result.workbook.activeTab).toBe(old.activeSheetIndex);
  });

  it.each(committed.filter((f) => f.expectImportError).map((f) => [f.id, f.filename]))('%s: either opens (damage is reported, not thrown) or refuses with an ImportError', async (_id, filename) => {
    let outcome: unknown;
    try {
      outcome = await readWorkbookParts(await XlsxPackage.open(corpus(filename)));
    } catch (e) {
      expect((e as Error).name).toBe('ImportError');
      return;
    }
    expect(outcome).toHaveProperty('workbook');
  });
});

describe('producer variety', () => {
  it.each(['openpyxl-image.xlsx', 'openpyxl-chart.xlsx', 'openpyxl-comment.xlsx'])('%s (absolute targets, which crash the exceljs reader) reads', async (file) => {
    const { workbook, warnings } = await parts(file);
    expect(workbook.sheets[0]).toMatchObject({ name: 'Data', state: 'visible', part: 'xl/worksheets/sheet1.xml' });
    expect(workbook.parts).toEqual({ styles: 'xl/styles.xml', theme: 'xl/theme/theme1.xml' });
    expect(warnings).toEqual([]);
  });
  it('openpyxl: locates the comment part and its VML drawing', async () => {
    const { workbook } = await parts('openpyxl-comment.xlsx');
    expect(workbook.sheets[0].commentParts).toEqual({ comments: 'xl/comments/comment1.xml', vmlDrawing: 'xl/drawings/commentsDrawing1.vml', threadedComments: [] });
    expect(workbook.calc.fullCalcOnLoad).toBe(true);
  });
  it('LibreOffice: hidden sheet, defined name, shared strings, lower-case type URIs, comments', async () => {
    const { workbook, sharedStrings } = await parts('libreoffice-rich.xlsx');
    expect(workbook.appName).toBe('Calc');
    expect(workbook.sheets.map((s) => [s.name, s.state])).toEqual([['Data', 'visible'], ['Hidden', 'hidden']]);
    expect(workbook.definedNames).toEqual([{ name: 'Quarter1', ref: 'Data!$B$2:$B$5', builtin: false, hidden: false }]);
    expect(workbook.sheets[0].commentParts.comments).toBe('xl/comments1.xml');
    expect(workbook.sheets[0].commentParts.vmlDrawing).toBe('xl/drawings/vmlDrawing1.vml');
    expect(workbook.parts.sharedStrings).toBe('xl/sharedStrings.xml');
    expect(sharedStrings.map((s) => s.text).slice(0, 3)).toEqual(['Region', 'Q1', 'Q2']);
    expect(sharedStrings).toHaveLength(10);
  });
  it('xlsxwriter: rich strings keep their runs and fonts', async () => {
    const { sharedStrings, workbook } = await parts('xlsxwriter-rich.xlsx');
    const rich = sharedStrings.filter((s) => s.runs);
    expect(rich.length).toBeGreaterThan(0);
    expect(rich[0].runs?.map((r) => r.text).join('')).toBe(rich[0].text);
    expect(rich[0].runs?.[1].font).toMatchObject({ bold: true, size: 11, name: 'Calibri', scheme: 'minor', color: { theme: 1, resolved: '#000000' } });
    expect(workbook.parts.sharedStrings).toBe('xl/sharedStrings.xml');
  });
  it('the hand-authored rich-text file keeps bold, italic and coloured runs', async () => {
    const { sharedStrings } = await parts('rich-text.xlsx');
    const fonts = sharedStrings.flatMap((s) => s.runs ?? []).map((r) => r.font ?? {});
    expect(fonts.some((f) => f.bold)).toBe(true);
    expect(fonts.some((f) => f.italic)).toBe(true);
    expect(fonts.some((f) => f.color?.resolved)).toBe(true);
  });
});

describe('workbook facts from the corpus', () => {
  it('dates-1904 is the 1904 system and the others are 1900', async () => {
    expect((await parts('dates-1904.xlsx')).workbook.date1904).toBe(true);
    expect((await parts('cell-types.xlsx')).workbook.date1904).toBe(false);
  });
  it('names.xlsx has defined names with scopes', async () => {
    const { definedNames } = (await parts('names.xlsx')).workbook;
    expect(definedNames.length).toBeGreaterThan(0);
    expect(definedNames.some((n) => n.localSheetId !== undefined)).toBe(true);
  });
  it('sheet-properties.xlsx has a non-visible sheet', async () => {
    expect((await parts('sheet-properties.xlsx')).workbook.sheets.some((s) => s.state !== 'visible')).toBe(true);
  });
  it('comments.xlsx locates notes and threaded comments with their persons part', async () => {
    const { workbook } = await parts('comments.xlsx');
    expect(workbook.sheets.some((s) => s.commentParts.comments !== undefined)).toBe(true);
    expect(workbook.sheets.some((s) => s.commentParts.threadedComments.length > 0)).toBe(true);
    expect(workbook.parts.persons).toBeDefined();
  });
  it('duplicate sheet names are all kept', async () => {
    expect((await parts('edge-duplicate-sheet-names.xlsx')).workbook.sheets.map((s) => s.name)).toEqual(['Data', 'Data', 'DATA']);
  });
  it('broken relationships: the dangling sheet stays listed, with warnings and fallbacks', async () => {
    const { workbook, warnings, theme, sharedStrings } = await parts('edge-bad-rels.xlsx');
    expect(workbook.sheets.map((s) => [s.name, s.part])).toEqual([['Present', 'xl/worksheets/sheet1.xml'], ['Dangling', undefined]]);
    expect(warnings).toEqual(['The sheet "Dangling" has no usable relationship (rId99).']);
    expect(workbook.parts.sharedStrings).toBe('xl/sharedStrings.xml'); // the relationship climbs out of the package: conventional name
    expect(theme.colors).toHaveLength(12);
    expect(sharedStrings.length).toBeGreaterThan(0);
  });
  it('invalid XML in a sheet part is not this layer\'s business: the workbook-level parts still read, without warnings', async () => {
    const result = await parts('edge-invalid-xml.xlsx');
    expect(result.workbook.sheets.map((s) => s.name)).toEqual(['Broken']);
    // the damage in this fixture is in the sheet (a stray element); the lenient X4 parser still reads every part, so the strings are all there and nothing warns
    expect(result.sharedStrings.map((s) => s.text)).toEqual(['a', 'b', 'c']);
    expect(result.warnings).toEqual([]);
  });
  it('XML entities (DOCTYPE with entity declarations) are inert', async () => {
    const { sharedStrings, warnings } = await parts('edge-xml-entities.xlsx');
    expect(JSON.stringify(sharedStrings)).not.toMatch(/root:|\/etc\/passwd/);
    expect(sharedStrings).toEqual([{ text: 'plain' }]);
    expect(warnings).toEqual([]);
  });
});

describe('theme, indexed and tint colours resolve (styles-theme.xlsx)', () => {
  it('gives every theme/indexed colour a resolved value, and tints differ from their base', async () => {
    const { styles, theme } = await parts('styles-theme.xlsx');
    expect(theme.colors).toHaveLength(12);
    const colours = [
      ...styles.fonts.map((f) => f.color),
      ...styles.fills.flatMap((f) => (f.type === 'pattern' ? [f.fg, f.bg] : f.stops.map((s) => s.color))),
    ].filter((c) => c !== undefined);
    const themed = colours.filter((c) => c.theme !== undefined || c.indexed !== undefined);
    expect(themed.length).toBeGreaterThan(0);
    for (const c of themed) expect(c.resolved).toMatch(/^#[0-9a-f]{6}$/);
    const tinted = themed.filter((c) => c.tint !== undefined && c.tint !== 0 && c.theme !== undefined);
    expect(tinted.length).toBeGreaterThan(0);
    for (const c of tinted) expect(c.resolved).not.toBe(theme.colors[c.theme as number]);
  });
  it('styles-basic and number-formats keep the number formats, fills and borders', async () => {
    const basic = (await parts('styles-basic.xlsx')).styles;
    expect(basic.cellXfs.length).toBeGreaterThan(1);
    expect(basic.borders.some((b) => b.left || b.top || b.bottom || b.right)).toBe(true);
    expect((await parts('number-formats.xlsx')).styles.numFmts.length).toBeGreaterThan(10);
  });
  it('conditional formatting workbooks carry dxfs', async () => {
    expect((await parts('cf-basic.xlsx')).styles.dxfs.length).toBeGreaterThan(0);
  });
});

describe('the internal API surface S2b builds on', () => {
  it('exports the part readers, the package and the colour helpers', () => {
    expect(Object.keys(api).sort()).toEqual([
      'DEFAULT_INDEXED_COLORS', 'DEFAULT_THEME_COLORS', 'PRESET_COLORS', 'XlsxPackage', 'applyTint', 'contentTypeOf', 'decodePart', 'decodeStringEscapes',
      'findWorkbookPart', 'forEachSharedString', 'isBinaryWorkbook', 'numberFormatCode', 'parseContentTypes', 'parseRelationships', 'parseRgb', 'parseSharedStrings',
      'parseStyles', 'parseTheme', 'partsWithContentType', 'readColor', 'readFont', 'readSharedString', 'readWorkbook', 'readWorkbookParts',
      'relsOfKind', 'relsPart', 'resolveColor', 'resolveRelationshipTarget', 'toImportExtras',
    ]);
    for (const value of Object.values(api)) expect(value).toBeDefined();
  });
});

describe('readWorkbookParts', () => {
  it('uses defaults for a workbook without styles, theme or strings', async () => {
    const result = await readWorkbookParts(await XlsxPackage.open(await workbookZip('<sheets/>', [])));
    expect(result.styles.cellXfs).toEqual([]);
    expect(result.sharedStrings).toEqual([]);
    expect(result.colors.themeColors).toHaveLength(12);
    expect(result.colors.indexedColors).toBeUndefined();
    expect(result.warnings).toEqual([]);
  });
  it('exposes the colour context built from the theme and the workbook palette', async () => {
    const result = await readWorkbookParts(
      await XlsxPackage.open(
        await workbookZip('<sheets/>', [['rId1', 'styles', 'styles.xml']], {
          'xl/styles.xml': '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><colors><indexedColors><rgbColor rgb="FF123456"/></indexedColors></colors></styleSheet>',
        }),
      ),
    );
    expect(result.colors.indexedColors).toEqual(['#123456']);
  });
});

describe('the reader output in the io result shape', () => {
  it('toImportExtras carries workbook, styles, shared strings (rich runs kept) and theme, structured-cloneable', async () => {
    const result = await parts('rich-text.xlsx');
    const extras = toImportExtras(result);
    expect(Object.keys(extras).sort()).toEqual(['sharedStrings', 'styles', 'theme', 'workbook']);
    expect(extras.sharedStrings).toBe(result.sharedStrings);
    expect(extras.theme.colors).toHaveLength(12);
    expect(extras.sharedStrings.some((s) => s.runs !== undefined && s.runs.length > 1)).toBe(true);
    const assignable: Pick<SpreadsheetImportResult, 'workbook' | 'styles' | 'sharedStrings' | 'theme'> = extras;
    expect(structuredClone(assignable)).toEqual(extras);
  });
  it('warnings is a snapshot: reading more parts afterwards does not change it', async () => {
    const pkg = await XlsxPackage.open(corpus('edge-bad-rels.xlsx'));
    const result = await readWorkbookParts(pkg);
    const before = [...result.warnings];
    await pkg.xml('xl/nothing-here.xml');
    pkg.warnings.push('later');
    expect(result.warnings).toEqual(before);
    expect(result.warnings).not.toBe(pkg.warnings);
  });
  it('a theme slot that cannot be read is a warning in the result', async () => {
    const result = await readWorkbookParts(
      await XlsxPackage.open(
        await workbookZip('<sheets/>', [['rId1', 'theme', 'theme/theme1.xml']], {
          'xl/theme/theme1.xml': '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"><a:dk1><a:schemeClr val="tx1"/></a:dk1></a:clrScheme></a:themeElements></a:theme>',
        }),
      ),
    );
    expect(result.warnings).toEqual(['The theme colour dk1 (schemeClr) could not be read; the Office default is used.']);
    expect(result.theme.colors[1]).toBe('#000000');
  });
});
