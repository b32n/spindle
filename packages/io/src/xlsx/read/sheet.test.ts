import type { SpreadsheetCell, SpreadsheetSheet, XlsxRelationship, XlsxSharedString, XlsxSheetRef } from '@b32nio/spindle-shared/export/xlsx';
import { DEFAULT_THEME_COLORS } from './colors';
import { isoToSerial, readSheet } from './sheet';
import { parseStyles } from './styles';
import { el, MAIN_NS, REL_NS } from './test-support';

const STYLES =
  '<fonts><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFF0000"/></font></fonts>' +
  '<fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor theme="5"/></patternFill></fill></fills>' +
  '<borders><border/></borders>' +
  '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0&quot; kg&quot;"/></numFmts>' +
  '<cellXfs><xf/><xf fontId="1" fillId="1"/><xf numFmtId="14"/><xf numFmtId="164"/><xf numFmtId="22"/></cellXfs>';

interface Options {
  head?: string;
  tail?: string;
  strings?: Array<string | XlsxSharedString>;
  date1904?: boolean;
  rels?: XlsxRelationship[];
  state?: XlsxSheetRef['state'];
  name?: string;
}

const refOf = (over: Partial<XlsxSheetRef> = {}): XlsxSheetRef => ({
  name: 'S',
  sheetId: 1,
  state: 'visible',
  relId: 'rId1',
  kind: 'worksheet',
  part: 'xl/worksheets/sheet1.xml',
  rels: [],
  commentParts: { threadedComments: [] },
  ...over,
});

function read(body: string, o: Options = {}): { sheet: SpreadsheetSheet; warnings: string[] } {
  const warnings: string[] = [];
  const ref = refOf({ name: o.name ?? 'S', state: o.state ?? 'visible', rels: o.rels ?? [] });
  const xml = `<worksheet ${MAIN_NS} ${REL_NS}>${o.head ?? ''}<sheetData>${body}</sheetData>${o.tail ?? ''}</worksheet>`;
  return { sheet: run(xml, ref, o, warnings), warnings };
}

function run(xml: string | undefined, ref: XlsxSheetRef, o: Options, warnings: string[]): SpreadsheetSheet {
  return readSheet(xml, ref, {
    styles: parseStyles(el(`<styleSheet ${MAIN_NS}>${STYLES}</styleSheet>`), DEFAULT_THEME_COLORS),
    sharedStrings: (o.strings ?? []).map((s) => (typeof s === 'string' ? { text: s } : s)),
    colors: { themeColors: DEFAULT_THEME_COLORS },
    date1904: o.date1904 ?? false,
    warn: (m) => warnings.push(m),
  });
}

const cell = (sheet: SpreadsheetSheet, key: string): SpreadsheetCell | undefined => sheet.cells.get(key);
const row = (r: number, cells: string, attrs = '') => `<row r="${r}"${attrs}>${cells}</row>`;

describe('readSheet: values', () => {
  it('reads numbers (the default type), strings from the table, formula strings, booleans and errors', () => {
    const { sheet, warnings } = read(
      row(1, '<c r="A1"><v>42.5</v></c><c r="B1" t="n"><v>-1e3</v></c><c r="C1" t="s"><v>1</v></c><c r="D1" t="str"><f>"x"&amp;"y"</f><v>xy</v></c><c r="E1" t="b"><v>1</v></c><c r="F1" t="b"><v>0</v></c><c r="G1" t="e"><v>#DIV/0!</v></c>'),
      { strings: ['zero', { text: 'one rich', runs: [{ text: 'one ' }, { text: 'rich' }] }] }
    );
    expect(warnings).toEqual([]);
    expect(cell(sheet, '0:0')).toEqual({ value: 42.5 });
    expect(cell(sheet, '0:1')).toEqual({ value: -1000 });
    expect(cell(sheet, '0:2')).toEqual({ value: 'one rich', sharedString: 1 });
    expect(cell(sheet, '0:3')).toEqual({ value: 'xy', formula: '"x"&"y"' });
    expect(cell(sheet, '0:4')).toEqual({ value: true });
    expect(cell(sheet, '0:5')).toEqual({ value: false });
    expect(cell(sheet, '0:6')).toEqual({ value: '#DIV/0!' });
  });

  it('reads inline strings, joining runs and leaving phonetic text out', () => {
    const { sheet } = read(
      row(1, '<c r="A1" t="inlineStr"><is><t>plain</t></is></c><c r="B1" t="inlineStr"><is><r><t>ru</t></r><r><rPr><b/></rPr><t>ns</t></r><rPh sb="0" eb="1"><t>ふり</t></rPh><phoneticPr fontId="0"/></is></c><c r="C1" t="inlineStr"><v>from v</v></c><c r="D1" t="inlineStr"/>')
    );
    expect(cell(sheet, '0:0')).toEqual({ value: 'plain' });
    expect(cell(sheet, '0:1')).toEqual({ value: 'runs' });
    expect(cell(sheet, '0:2')).toEqual({ value: 'from v' });
    expect(cell(sheet, '0:3')).toBeUndefined();
  });

  it('decodes entities and CDATA in values', () => {
    const { sheet } = read(row(1, '<c r="A1" t="inlineStr"><is><t xml:space="preserve"> a &amp; b &lt;<![CDATA[<c>]]></t></is></c>'));
    expect(cell(sheet, '0:0')?.value).toBe(' a & b <<c>');
  });

  it('reads ISO dates (t="d") as serial numbers', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1" t="d"><v>1900-01-01</v></c><c r="B1" t="d"><v>2024-03-15T12:00:00</v></c><c r="C1" t="d"><v>nonsense</v></c><c r="D1" t="d"/>'));
    expect(cell(sheet, '0:0')?.value).toBe(2);
    expect(cell(sheet, '0:1')?.value).toBe(45366.5);
    expect(cell(sheet, '0:2')?.value).toBe('nonsense');
    expect(warnings).toEqual(['S!C1: "nonsense" is not a date; imported as text']);
    expect(cell(sheet, '0:3')).toBeUndefined();
  });

  it('does not keep an empty cell, and keeps a formula without a result', () => {
    const { sheet } = read(
      row(1, '<c r="A1"/><c r="B1" s="1"/><c r="C1"><v/></c><c r="D1"><v>  </v></c><c r="E1" t="s"><v></v></c><c r="F1"><f>A1+1</f></c><c r="G1" t="str"><f>"x"</f></c><c r="H1" t="b"><f>TRUE</f></c><c r="I1" t="e"><f>NA()</f></c><c r="J1" t="d"><f>NOW()</f></c>')
    );
    expect([...sheet.cells.keys()]).toEqual(['0:5', '0:6', '0:7', '0:8', '0:9']);
    expect(cell(sheet, '0:5')).toEqual({ value: null, formula: 'A1+1' });
    expect(cell(sheet, '0:9')).toEqual({ value: null, formula: 'NOW()' });
  });

  it('warns about text in a number cell and keeps it as text', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1"><v>12abc</v></c>'));
    expect(cell(sheet, '0:0')?.value).toBe('12abc');
    expect(warnings).toEqual(['S!A1: "12abc" is not a number; imported as text']);
  });

  it('treats a shared string index the table does not have as an empty cell, with one warning per sheet', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>7</v></c><c r="C1" t="s"><v>-1</v></c><c r="D1" t="s"><v>x</v></c><c r="E1" t="s"><v>1.5</v></c><c r="F1" t="s"><v>__proto__</v></c><c r="G1" t="s"><v>constructor</v></c>'), { strings: ['ok'] });
    expect([...sheet.cells.keys()]).toEqual(['0:0']);
    expect(warnings).toEqual(['S: 6 cell(s) refer to a shared string the file does not have; imported as empty.']);
  });

  it('moves 1904-system dates to 1900 serials, only where the number format shows a date', () => {
    const { sheet } = read(row(1, '<c r="A1" s="2"><v>0</v></c><c r="B1"><v>5</v></c><c r="C1" s="3"><v>5</v></c><c r="D1" s="2"><f>NOW()</f><v>10</v></c><c r="E1" s="2" t="str"><v>text</v></c>'), { date1904: true });
    expect(cell(sheet, '0:0')?.value).toBe(1462);
    expect(cell(sheet, '0:1')?.value).toBe(5);
    expect(cell(sheet, '0:2')?.value).toBe(5);
    expect(cell(sheet, '0:3')?.value).toBe(1472);
    expect(cell(sheet, '0:4')?.value).toBe('text');
    expect(cell(read(row(1, '<c r="A1" s="2"><v>0</v></c>')).sheet, '0:0')?.value).toBe(0);
  });
});

describe('isoToSerial', () => {
  it.each([
    ['1899-12-30', 0],
    ['1900-01-01', 2],
    ['2000-01-01', 36526],
    ['2000-01-01T06:00', 36526.25],
    ['2000-01-01T06:00:30Z', 36526 + 6 / 24 + 30 / 86400],
    ['2000-01-01T00:00:01.5', 36526 + 1.5 / 86400],
  ])('%s', (text, serial) => {
    expect(isoToSerial(text)).toBeCloseTo(serial, 9);
  });

  it.each(['', 'x', '2000-1-1', '2000-01-01T6:00', '01/01/2000', '2000-13-45x'])('%j is not a date', (text) => {
    expect(isoToSerial(text)).toBeUndefined();
  });
});

describe('readSheet: cell addresses', () => {
  it('follows r, and counts on from the previous cell or row when r is missing', () => {
    const { sheet } = read('<row><c><v>1</v></c><c><v>2</v></c><c r="D1"><v>3</v></c><c><v>4</v></c></row><row><c><v>5</v></c></row><row r="10"><c r="B10"><v>6</v></c></row><row><c><v>7</v></c></row>');
    expect([...sheet.cells.entries()].map(([k, c]) => [k, c.value])).toEqual([
      ['0:0', 1], ['0:1', 2], ['0:3', 3], ['0:4', 4], ['1:0', 5], ['9:1', 6], ['10:0', 7],
    ]);
  });

  it('lets the cell address win over the row, and counts on past a bad address', () => {
    const { sheet } = read('<row r="3"><c r="A5"><v>1</v></c><c r="oops"><v>2</v></c><c r="A0"><v>3</v></c></row>');
    expect([...sheet.cells.keys()]).toEqual(['4:0', '4:1', '4:2']);
  });

  it('reads a row r of zero, a bad r and a huge r as "the next row"', () => {
    const { sheet } = read('<row r="0"><c><v>1</v></c></row><row r="x"><c><v>2</v></c></row>');
    expect([...sheet.cells.keys()]).toEqual(['0:0', '1:0']);
  });

  it('reads prefixed markup like plain markup', () => {
    const xml = `<x:worksheet xmlns:x="${'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}"><x:sheetData><x:row r="2"><x:c r="B2" t="s"><x:v>0</x:v></x:c></x:row></x:sheetData></x:worksheet>`;
    const sheet = run(xml, refOf({ name: 'P' }), { strings: ['hi'] }, []);
    expect(cell(sheet, '1:1')).toEqual({ value: 'hi', sharedString: 0 });
  });

  it('treats attribute and key names that are Object.prototype members as ordinary', () => {
    const { sheet, warnings } = read('<row __proto__="1" constructor="2" r="1"><c __proto__="x" constructor="y" toString="z" r="A1"><v>1</v></c></row>');
    expect(cell(sheet, '0:0')).toEqual({ value: 1 });
    expect(warnings).toEqual([]);
    expect(({} as Record<string, unknown>).polluted).toBeUndefined();
  });
});

describe('readSheet: hostile names in the small elements', () => {
  it('reads elements whose attributes are named like Object.prototype members', () => {
    const rels: XlsxRelationship[] = [{ id: 'rId1', type: 'x/hyperlink', kind: 'hyperlink', target: 'https://example.com/', external: true }];
    const { sheet, warnings } = read(row(1, '<c r="A1"><v>1</v></c>'), {
      rels,
      head: '<sheetViews><sheetView __proto__="1" constructor="2"><pane __proto__="x" toString="y" ySplit="1" state="frozen"/></sheetView></sheetViews><cols><col __proto__="x" constructor="y" valueOf="z" min="1" max="1" width="5"/></cols>',
      tail: '<hyperlinks><hyperlink hasOwnProperty="x" __proto__="y" ref="A1" r:id="rId1"/></hyperlinks><mergeCells><mergeCell constructor="x" ref="A1:B1"/></mergeCells>',
    });
    expect(warnings).toEqual([]);
    expect(sheet.frozenRows).toBe(1);
    expect(sheet.columnWidths?.get(0)).toBe(40);
    expect(cell(sheet, '0:0')?.hyperlink).toBe('https://example.com/');
    expect(sheet.mergedRanges).toHaveLength(1);
    expect(Object.keys(Object.prototype)).toEqual([]);
  });
});

describe('readSheet: formulas', () => {
  it('reads the text of a plain formula and ignores an empty one', () => {
    const { sheet } = read(row(1, '<c r="A1"><f>SUM(B1:B2)</f><v>3</v></c><c r="B1"><f></f><v>1</v></c><c r="C1"><f/><v>2</v></c>'));
    expect(cell(sheet, '0:0')).toEqual({ value: 3, formula: 'SUM(B1:B2)' });
    expect(cell(sheet, '0:1')).toEqual({ value: 1 });
    expect(cell(sheet, '0:2')).toEqual({ value: 2 });
  });

  it('expands shared formulas: each cell of the group gets the master formula moved by its distance', () => {
    const { sheet, warnings } = read(
      row(1, '<c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2+$A$1</f><v>2</v></c>') +
        row(2, '<c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c><c r="C2"><f t="shared" si="0"/><v>9</v></c>') +
        row(3, '<c r="B3"><f t="shared" si="0"/><v>6</v></c><c r="D3"><f t="shared" ref="D3:D4" si="1">B3</f><v>0</v></c>') +
        row(4, '<c r="D4"><f t="shared" si="1"></f><v>0</v></c>')
    );
    expect(warnings).toEqual([]);
    expect(cell(sheet, '0:1')).toEqual({ value: 2, formula: 'A1*2+$A$1', formulaInfo: { type: 'shared', si: 0, ref: 'B1:B3' } });
    expect(cell(sheet, '1:1')).toEqual({ value: 4, formula: 'A2*2+$A$1', formulaInfo: { type: 'shared', si: 0 } });
    expect(cell(sheet, '1:2')?.formula).toBe('B2*2+$A$1');
    expect(cell(sheet, '2:1')?.formula).toBe('A3*2+$A$1');
    expect(cell(sheet, '3:3')?.formula).toBe('B4');
  });

  it('keeps a shared master without a ref, and expands a group whose master comes later in the file as unresolved', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1"><f t="shared" si="3"/><v>1</v></c><c r="B1"><f t="shared" si="3">X1</f><v>2</v></c><c r="C1"><f t="shared" si="3"/><v>3</v></c>'));
    expect(cell(sheet, '0:0')).toEqual({ value: 1 });
    expect(cell(sheet, '0:1')?.formulaInfo).toEqual({ type: 'shared', si: 3 });
    expect(cell(sheet, '0:2')?.formula).toBe('Y1');
    expect(warnings).toEqual(['S!A1: shared formula without a resolvable expression; imported as a plain value']);
  });

  it('reads a shared formula with a missing or unusable si as a plain formula', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1"><f t="shared">B1</f><v>1</v></c><c r="B1"><f t="shared" si="__proto__">C1</f><v>1</v></c><c r="C1"><f t="shared" si="1.5"/><v>1</v></c>'));
    expect(cell(sheet, '0:0')).toEqual({ value: 1, formula: 'B1' });
    expect(cell(sheet, '0:1')).toEqual({ value: 1, formula: 'C1' });
    expect(cell(sheet, '0:2')).toEqual({ value: 1 });
    expect(warnings).toEqual([]);
  });

  it('keeps the text and the range of an array formula on its cell and leaves the cells it covers as values', () => {
    const { sheet } = read(row(1, '<c r="A1"><f t="array" ref="A1:A3">ROW(1:3)</f><v>1</v></c>') + row(2, '<c r="A2"><v>2</v></c>') + row(3, '<c r="A3"><v>3</v></c><c r="B3"><f t="array">SUM(A1:A3)</f><v>6</v></c><c r="C3"><f t="array" ref="C3"/><v>0</v></c>'));
    expect(cell(sheet, '0:0')).toEqual({ value: 1, formula: 'ROW(1:3)', formulaInfo: { type: 'array', ref: 'A1:A3' } });
    expect(cell(sheet, '1:0')).toEqual({ value: 2 });
    expect(cell(sheet, '2:1')).toEqual({ value: 6, formula: 'SUM(A1:A3)', formulaInfo: { type: 'array' } });
    expect(cell(sheet, '2:2')).toEqual({ value: 0, formulaInfo: { type: 'array', ref: 'C3' } });
  });

  it('notes a data table without making a formula of it', () => {
    const { sheet } = read(row(1, '<c r="A1"><f t="dataTable" ref="A1:B2" dt2D="1" r1="X1" r2="Y1"/><v>5</v></c><c r="B1"><f t="dataTable"/><v>6</v></c>'));
    expect(cell(sheet, '0:0')).toEqual({ value: 5, formulaInfo: { type: 'dataTable', ref: 'A1:B2' } });
    expect(cell(sheet, '0:1')).toEqual({ value: 6, formulaInfo: { type: 'dataTable' } });
  });
});

describe('readSheet: formats', () => {
  it('gives each cell its format: style, number format and the format index; the default format adds nothing', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1"><v>1</v></c><c r="B1" s="0"><v>1</v></c><c r="C1" s="1"><v>1</v></c><c r="D1" s="3"><v>1</v></c>'));
    expect(warnings).toEqual([]);
    expect(cell(sheet, '0:0')).toEqual({ value: 1 });
    expect(cell(sheet, '0:1')).toEqual({ value: 1, styleIndex: 0 });
    expect(cell(sheet, '0:2')).toEqual({ value: 1, styleIndex: 1, style: { bold: true, fontColor: '#ff0000', backgroundColor: '#ed7d31' } });
    expect(cell(sheet, '0:3')).toEqual({ value: 1, styleIndex: 3, style: { fontFamily: 'Calibri', fontSize: 11 }, numberFormat: { code: '0.0" kg"' } });
  });

  it('shares one style object between the cells of a format', () => {
    const { sheet } = read(row(1, '<c r="A1" s="1"><v>1</v></c><c r="B1" s="1"><v>2</v></c>'));
    expect(cell(sheet, '0:0')?.style).toBe(cell(sheet, '0:1')?.style);
  });

  it('warns once about formats the file does not have, and reads the cell unformatted', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1" s="99"><v>1</v></c><c r="B1" s="-1"><v>2</v></c><c r="C1" s="x"><v>3</v></c><c r="D1" s="1.5"><v>4</v></c><c r="E1" s="__proto__"><v>5</v></c>'));
    expect([...sheet.cells.values()]).toEqual([{ value: 1 }, { value: 2 }, { value: 3 }, { value: 4 }, { value: 5 }]);
    expect(warnings).toEqual(['S: 5 cell(s) use a format the file does not have; imported unformatted.']);
  });
});

describe('readSheet: merges', () => {
  it('reads the ranges and drops the cells a merge covers but its first', () => {
    const { sheet } = read(row(1, '<c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c>') + row(2, '<c r="A2"><v>4</v></c><c r="B2"><v>5</v></c><c r="D2"><v>6</v></c>'), {
      tail: '<mergeCells count="3"><mergeCell ref="A1:B2"/><mergeCell ref="bad"/><mergeCell/><mergeCell ref="D2:C1"/></mergeCells>',
    });
    expect(sheet.mergedRanges).toEqual([
      { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      { startRow: 0, startCol: 2, endRow: 1, endCol: 3 },
    ]);
    expect([...sheet.cells.keys()]).toEqual(['0:0', '0:2']);
  });

  it('walks the cells instead of the range for a huge merge', () => {
    const { sheet } = read(row(1, '<c r="A1"><v>1</v></c><c r="C3"><v>2</v></c><c r="XFD1048576"><v>3</v></c>'), {
      tail: '<mergeCells><mergeCell ref="A1:XFD1048576"/></mergeCells>',
    });
    expect([...sheet.cells.keys()]).toEqual(['0:0']);
    expect(sheet.mergedRanges).toHaveLength(1);
  });

  it('recomputes the extent from what is left: widths past the last kept cell are not read', () => {
    const { sheet } = read(row(1, '<c r="A1"><v>1</v></c><c r="Z1"><v>2</v></c>'), {
      head: '<cols><col min="1" max="30" width="10"/></cols>',
      tail: '<mergeCells><mergeCell ref="A1:Z1"/><mergeCell ref="A1:B1"/></mergeCells>',
    });
    expect([...sheet.cells.keys()]).toEqual(['0:0']);
    expect([...(sheet.columnWidths as Map<number, number>).keys()]).toEqual([0]);
  });

  it('has no mergedRanges without merges', () => {
    expect(read(row(1, '<c r="A1"><v>1</v></c>')).sheet.mergedRanges).toBeUndefined();
  });
});

describe('readSheet: hyperlinks', () => {
  const rels: XlsxRelationship[] = [{ id: 'rId1', type: 'x/hyperlink', kind: 'hyperlink', target: 'https://example.com/a', external: true }];

  it('links a cell by its relationship target, keeping its value and formula', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1" t="s"><v>0</v></c><c r="B1"><f>A1&amp;"!"</f><v>1</v></c><c r="C1"><v>3</v></c>'), {
      strings: ['click'],
      rels,
      tail: '<hyperlinks><hyperlink ref="A1" r:id="rId1" display="d" tooltip="t"/><hyperlink ref="B1" r:id="rId1"/></hyperlinks>',
    });
    expect(warnings).toEqual([]);
    expect(cell(sheet, '0:0')).toMatchObject({ value: 'click', hyperlink: 'https://example.com/a' });
    expect(cell(sheet, '0:1')).toMatchObject({ value: 1, formula: 'A1&"!"', hyperlink: 'https://example.com/a' });
    expect(cell(sheet, '0:2')?.hyperlink).toBeUndefined();
    expect(sheet.details?.hyperlinks).toEqual([
      { ref: 'A1', target: 'https://example.com/a', display: 'd', tooltip: 't' },
      { ref: 'B1', target: 'https://example.com/a' },
    ]);
  });

  it('links every kept cell of a range, whichever of range and cells is smaller', () => {
    const body = row(1, '<c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c>');
    const small = read(body, { rels, tail: '<hyperlinks><hyperlink ref="A1:B2" r:id="rId1"/></hyperlinks>' }).sheet;
    expect([...small.cells.values()].map((c) => c.hyperlink)).toEqual(['https://example.com/a', 'https://example.com/a', undefined]);
    const huge = read(body, { rels, tail: '<hyperlinks><hyperlink ref="A1:XFD100" r:id="rId1"/></hyperlinks>' }).sheet;
    expect([...huge.cells.values()].map((c) => c.hyperlink)).toEqual(Array(3).fill('https://example.com/a'));
  });

  it('keeps an internal link (location) in details only, and ignores a link with no usable reference', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1"><v>1</v></c>'), {
      rels,
      tail: '<hyperlinks><hyperlink ref="A1" location="Sheet2!A1"/><hyperlink ref="nonsense" r:id="rId1"/><hyperlink r:id="rId1"/></hyperlinks>',
    });
    expect(cell(sheet, '0:0')?.hyperlink).toBeUndefined();
    expect(sheet.details?.hyperlinks).toEqual([{ ref: 'A1', location: 'Sheet2!A1' }, { ref: 'nonsense', target: 'https://example.com/a' }]);
    expect(warnings).toEqual([]);
  });

  it('warns about a link whose relationship the sheet does not have', () => {
    const { sheet, warnings } = read(row(1, '<c r="A1"><v>1</v></c>'), { tail: '<hyperlinks><hyperlink ref="A1" r:id="rId9"/></hyperlinks>' });
    expect(cell(sheet, '0:0')?.hyperlink).toBeUndefined();
    expect(warnings).toEqual(['The hyperlink at A1 names the relationship rId9, which the sheet does not have.']);
  });

  it('does not let a relationship id like __proto__ find anything', () => {
    const { warnings } = read(row(1, '<c r="A1"><v>1</v></c>'), { rels, tail: '<hyperlinks><hyperlink ref="A1" r:id="__proto__"/><hyperlink ref="A1" r:id="constructor"/></hyperlinks>' });
    expect(warnings).toHaveLength(2);
  });
});

describe('readSheet: columns and rows', () => {
  it('reads widths and hidden columns, clipped to the last column with a cell', () => {
    const { sheet } = read(row(1, '<c r="A1"><v>1</v></c><c r="D1"><v>1</v></c>'), {
      head:
        '<cols><col min="1" max="2" width="10"/><col min="3" max="3" hidden="1" width="20" customWidth="1"/><col min="2" max="2" hidden="0"/>' +
        '<col min="4" max="16384" width="30" hidden="true"/><col min="0" max="1" width="5"/><col min="5" max="3" width="5"/><col min="x" max="3"/><col max="3"/></cols>',
    });
    expect([...(sheet.columnWidths as Map<number, number>)]).toEqual([[0, 75], [1, 75], [2, 145], [3, 215]]);
    expect([...(sheet.hiddenCols as Set<number>)]).toEqual([2, 3]);
    expect(sheet.details?.columns).toHaveLength(4);
  });

  it('un-hides a column a later entry shows', () => {
    const { sheet } = read(row(1, '<c r="C1"><v>1</v></c>'), { head: '<cols><col min="1" max="3" hidden="1"/><col min="2" max="2" hidden="0"/></cols>' });
    expect([...(sheet.hiddenCols as Set<number>)]).toEqual([0, 2]);
  });

  it('reads every column attribute into details', () => {
    const { sheet } = read('', { head: '<cols><col min="2" max="4" width="12.5" hidden="1" customWidth="1" bestFit="1" style="3" outlineLevel="2" collapsed="1"/><col min="9" max="9"/></cols>' });
    expect(sheet.details?.columns).toEqual([
      { min: 2, max: 4, width: 12.5, hidden: true, customWidth: true, bestFit: true, style: 3, outlineLevel: 2, collapsed: true },
      { min: 9, max: 9 },
    ]);
  });

  it('reads row heights (points to px) and hidden rows up to the last row with a cell, and every row attribute into details', () => {
    const { sheet } = read(
      row(1, '<c r="A1"><v>1</v></c>', ' ht="15" customHeight="1" hidden="1" s="2" customFormat="1" outlineLevel="1" collapsed="0" spans="1:1"') +
        row(2, '<c r="A2"><v>1</v></c>', ' ht="30"') +
        row(3, '', ' hidden="1"') +
        row(4, '', ' ht="45"') +
        row(5, '<c r="A5"><v>1</v></c>') +
        row(9, '', ' ht="60" hidden="1"')
    );
    expect([...(sheet.rowHeights as Map<number, number>)]).toEqual([[0, 20], [1, 40], [3, 60]]);
    expect([...(sheet.hiddenRows as Set<number>)]).toEqual([0, 2]);
    expect(sheet.details?.rows).toEqual([
      { row: 0, height: 15, hidden: true, customHeight: true, style: 2, customFormat: true, outlineLevel: 1, collapsed: false },
      { row: 1, height: 30 },
      { row: 2, hidden: true },
      { row: 3, height: 45 },
      { row: 8, height: 60, hidden: true },
    ]);
  });

  it('has empty tables for a sheet with neither', () => {
    const { sheet } = read(row(1, '<c r="A1"><v>1</v></c>'));
    expect(sheet.columnWidths?.size).toBe(0);
    expect(sheet.rowHeights?.size).toBe(0);
    expect(sheet.details?.columns).toEqual([]);
    expect(sheet.details?.rows).toEqual([]);
  });
});

describe('readSheet: views, properties, defaults, extensions', () => {
  it('reads frozen panes', () => {
    const view = (pane: string) => read('', { head: `<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>` }).sheet;
    const a = view('<pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>');
    expect([a.frozenRows, a.frozenCols]).toEqual([3, 2]);
    expect(view('<pane ySplit="1" state="frozen"/>').frozenCols).toBeUndefined();
    expect(view('<pane xSplit="1" state="frozen"/>').frozenRows).toBeUndefined();
    expect(view('<pane xSplit="0" ySplit="0" state="frozen"/>').frozenRows).toBeUndefined();
    expect(view('<pane xSplit="2" ySplit="3" state="split"/>').frozenRows).toBeUndefined();
    expect(view('<pane xSplit="2" ySplit="3" state="frozenSplit"/>').frozenRows).toBeUndefined();
    expect(view('').frozenRows).toBeUndefined();
  });

  it('finds a frozen pane in any of several views', () => {
    const { sheet } = read('', { head: '<sheetViews><sheetView/><sheetView><pane ySplit="2" state="frozen"/></sheetView></sheetViews>' });
    expect(sheet.frozenRows).toBe(2);
    expect(sheet.details?.views).toHaveLength(2);
  });

  it('reads view attributes, the pane and the active cells into details', () => {
    const { sheet } = read('', {
      head:
        '<sheetViews><sheetView showGridLines="0" showRowColHeaders="0" showZeros="0" rightToLeft="1" tabSelected="1" zoomScale="150" zoomScaleNormal="100" view="pageBreakPreview" topLeftCell="B2">' +
        '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/><selection activeCell="C5" sqref="C5"/><selection pane="topRight"/><selection activeCell="A1"/></sheetView><sheetView/></sheetViews>',
    });
    expect(sheet.details?.views).toEqual([
      {
        showGridLines: false, showRowColHeaders: false, showZeros: false, rightToLeft: true, tabSelected: true, zoomScale: 150, zoomScaleNormal: 100, view: 'pageBreakPreview', topLeftCell: 'B2',
        pane: { xSplit: 1, ySplit: 2, topLeftCell: 'B3', activePane: 'bottomRight', state: 'frozen' },
        activeCells: ['C5', 'A1'],
      },
      { activeCells: [] },
    ]);
  });

  it('reads sheet properties, with the tab colour resolved', () => {
    const { sheet } = read('', { head: '<sheetPr codeName="Sheet1" filterMode="1"><tabColor theme="5" tint="0.4"/><outlinePr summaryBelow="0" summaryRight="1"/><pageSetUpPr fitToPage="1"/></sheetPr>' });
    expect(sheet.details?.properties).toEqual({
      codeName: 'Sheet1',
      filterMode: true,
      tabColor: { theme: 5, tint: 0.4, resolved: '#f4b183' },
      summaryBelow: false,
      summaryRight: true,
      fitToPage: true,
    });
    expect(read('', { head: '<sheetPr><outlinePr/></sheetPr>' }).sheet.details?.properties).toEqual({});
    expect(read('').sheet.details?.properties).toEqual({});
  });

  it('reads a self-closing sheetPr and dimension', () => {
    const { sheet } = read('', { head: '<sheetPr codeName="X"/><dimension ref="A1:C9"/>' });
    expect(sheet.details?.properties).toEqual({ codeName: 'X' });
    expect(sheet.details?.dimension).toBe('A1:C9');
  });

  it('reads the format defaults', () => {
    const { sheet } = read('', { head: '<sheetFormatPr baseColWidth="8" defaultColWidth="9.5" defaultRowHeight="15" customHeight="1" zeroHeight="1" outlineLevelRow="2" outlineLevelCol="3"/>' });
    expect(sheet.details?.format).toEqual({ baseColWidth: 8, defaultColWidth: 9.5, defaultRowHeight: 15, customHeight: true, zeroHeight: true, outlineLevelRow: 2, outlineLevelCol: 3 });
    expect(read('').sheet.details?.format).toEqual({});
  });

  it('captures extLst raw, as written, and the state and part of the sheet', () => {
    const ext = '<extLst><ext uri="{X}" xmlns:x14="urn:x14"><x14:sparklineGroups><x14:sparklineGroup/></x14:sparklineGroups></ext></extLst>';
    const { sheet } = read('', { tail: ext, state: 'hidden' });
    expect(sheet.details?.extLst).toBe(ext);
    expect(sheet.details?.state).toBe('hidden');
    expect(sheet.details?.part).toBe('xl/worksheets/sheet1.xml');
    expect(sheet.details?.wellFormed).toBe(true);
    expect(read('').sheet.details?.extLst).toBeUndefined();
    expect(read('').sheet.details?.dimension).toBeUndefined();
  });

  it('gives the sheet name as given and an empty part when the reference has none', () => {
    const noPart = run('<worksheet><sheetData/></worksheet>', refOf({ name: 'N', part: undefined }), {}, []);
    expect(noPart.name).toBe('N');
    expect(noPart.details?.part).toBe('');
  });
});

describe('readSheet: damage', () => {
  it('reads an absent part as an empty sheet', () => {
    const sheet = run(undefined, refOf({ name: 'Gone' }), {}, []);
    expect(sheet.cells.size).toBe(0);
    expect(sheet.details?.wellFormed).toBe(true);
  });

  it('keeps what precedes the damage, drops the cell the damage cut off, and says so', () => {
    const warnings: string[] = [];
    const xml = `<worksheet ${MAIN_NS}><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></row></sheetData></worksheet>`;
    const sheet = run(xml, refOf(), {}, warnings);
    expect([...sheet.cells.keys()]).toEqual(['0:0', '0:1', '1:0']);
    expect(sheet.details?.wellFormed).toBe(false);
    expect(warnings).toEqual(['The sheet "S" is not well-formed XML; it was read as far as the damage (3 cells).']);
  });

  it('reads nothing from a sheet nested past the depth cap', () => {
    const warnings: string[] = [];
    const sheet = run('<a>'.repeat(300), refOf({ name: 'D' }), {}, warnings);
    expect(sheet.cells.size).toBe(0);
    expect(warnings).toHaveLength(1);
  });

  it('limits the single-cell warnings of a sheet and counts the rest', () => {
    const cells = Array.from({ length: 25 }, (_, i) => `<c r="A${i + 1}"><v>x${i}</v></c>`).join('');
    const { warnings } = read(row(1, cells));
    expect(warnings).toHaveLength(21);
    expect(warnings[20]).toBe('S: 5 more cell warnings were left out.');
  });

  it('ignores small elements that are not well-formed on their own, and unknown elements', () => {
    const { sheet } = read(row(1, '<c r="A1"><v>1</v></c>'), { head: '<cols><col min="1" max="1" width="5"/></cols><unknownThing a="1"><c r="Z9"><v>9</v></c></unknownThing>', tail: '<mergeCells><mergeCell ref="B2:C3"/></mergeCells><pageMargins left="1"/>' });
    expect([...sheet.cells.keys()]).toEqual(['0:0']);
    expect(sheet.mergedRanges).toHaveLength(1);
  });

  it('keeps the first of a small element that is written twice', () => {
    const { sheet } = read('', { head: '<dimension ref="A1"/><dimension ref="B2"/><mergeCells><mergeCell ref="A1:B1"/></mergeCells><mergeCells><mergeCell ref="A2:B2"/></mergeCells>' });
    expect(sheet.details?.dimension).toBe('A1');
    expect(sheet.mergedRanges).toEqual([{ startRow: 0, startCol: 0, endRow: 0, endCol: 1 }]);
  });
});
