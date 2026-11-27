import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import { applicableParts, parseLook } from './table-props';
import { parseXml, find, childrenOf } from './xml-tree';
import type { DocBlock, DocSection, DocTableStyle } from '@b32nio/spindle-shared/export/docx';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
type Table = Extract<DocBlock, { type: 'table' }>;

async function docx(body: string, styles = ''): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types/>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  if (styles) zip.file('word/styles.xml', `<?xml version="1.0"?><w:styles ${NS}>${styles}</w:styles>`);
  return zip.generateAsync({ type: 'uint8array' });
}
const cell = (text: string, tcPr = '') => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (cells: string, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells}</w:tr>`;
const tbl = (rows: string, tblPr = '', grid = '') => `<w:tbl><w:tblPr>${tblPr}</w:tblPr>${grid ? `<w:tblGrid>${grid}</w:tblGrid>` : ''}${rows}</w:tbl>`;
const importTable = async (body: string, styles = ''): Promise<Table> => (await importDocx(await docx(body, styles))).sections[0].blocks[0] as Table;

describe('table formatting import', () => {
  it('reads table and cell borders (line style, width in px, colour), "nil" as none', async () => {
    const t = await importTable(
      tbl(
        row(cell('a', '<w:tcBorders><w:bottom w:val="double" w:sz="12" w:color="FF0000"/><w:left w:val="nil"/></w:tcBorders>')),
        '<w:tblBorders><w:top w:val="single" w:sz="8" w:color="auto"/><w:insideH w:val="dashed" w:sz="4" w:color="00FF00"/></w:tblBorders>'
      )
    );
    expect(t.borders).toEqual({ top: { style: 'single', widthPx: 1.33 }, insideH: { style: 'dashed', color: '#00ff00' } });
    expect(t.rows[0].cells[0].borders).toEqual({ bottom: { style: 'double', widthPx: 2, color: '#ff0000' }, left: { style: 'none' } });
  });

  it('does not keep table sides set to none, so they cancel a style border', async () => {
    const styles = '<w:style w:type="table" w:styleId="G"><w:name w:val="G"/><w:tblPr><w:tblBorders><w:top w:val="single"/><w:left w:val="single"/></w:tblBorders></w:tblPr></w:style>';
    const t = await importTable(tbl(row(cell('a')), '<w:tblStyle w:val="G"/><w:tblBorders><w:top w:val="nil"/></w:tblBorders>'), styles);
    expect(t.borders).toEqual({ left: { style: 'single' } });
  });

  it('gives a table with no border lines `borders: {}` (no lines), not an absent value', async () => {
    expect((await importTable(tbl(row(cell('a'))))).borders).toEqual({});
  });

  it('reads cell shading, vertical alignment, header rows and cantSplit', async () => {
    const t = await importTable(
      tbl(
        row(cell('h', '<w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/><w:vAlign w:val="center"/>'), '<w:tblHeader/><w:cantSplit/>') +
          row(cell('b', '<w:shd w:val="clear" w:fill="auto"/>'), '<w:tblHeader w:val="0"/>')
      )
    );
    expect(t.rows[0]).toMatchObject({ repeatHeader: true, cantSplit: true });
    expect(t.rows[0].cells[0]).toMatchObject({ backgroundColor: '#ddeeff', verticalAlign: 'center' });
    expect(t.rows[1].repeatHeader).toBeUndefined();
    expect(t.rows[1].cells[0].backgroundColor).toBeUndefined();
  });

  it('reads the grid as column widths and gives each cell the widths of the columns it spans', async () => {
    const t = await importTable(
      tbl(
        row(cell('a', '<w:gridSpan w:val="2"/>') + cell('b')) + row(cell('c') + cell('d') + cell('e')),
        '',
        '<w:gridCol w:w="1440"/><w:gridCol w:w="2880"/><w:gridCol w:w="1440"/>'
      )
    );
    expect(t.colWidthsPx).toEqual([96, 192, 96]);
    expect(t.rows[0].cells[0].colwidthPx).toEqual([96, 192]);
    expect(t.rows[0].cells[1].colwidthPx).toEqual([96]);
  });

  it('falls back to the cell width (tcW) when there is no grid', async () => {
    const t = await importTable(tbl(row(cell('a', '<w:tcW w:w="2880" w:type="dxa"/>'))));
    expect(t.colWidthsPx).toBeUndefined();
    expect(t.rows[0].cells[0].colwidthPx).toEqual([192]);
  });

  it('keeps the table style id and tblLook', async () => {
    const styles = '<w:style w:type="table" w:styleId="Fancy"><w:name w:val="Fancy"/></w:style>';
    const t = await importTable(tbl(row(cell('a')), '<w:tblStyle w:val="Fancy"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'), styles);
    expect(t.namedStyleId).toBe('Fancy');
    expect(t.look).toEqual({ firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true });
  });

  it('reads tblLook from the hex mask when the attributes are absent', () => {
    const root = parseXml(`<w:tblPr ${NS}><w:tblLook w:val="04A0"/></w:tblPr>`);
    const look = parseLook(find(childrenOf(root[0]), 'w:tblLook'));
    expect(look).toEqual({ firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true });
  });
});

const STYLE = `<w:style w:type="table" w:styleId="Banded"><w:name w:val="Banded"/>
  <w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4" w:color="999999"/></w:tblBorders></w:tblPr>
  <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="336699"/></w:tcPr></w:tblStylePr>
  <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr></w:tblStylePr>
  <w:tblStylePr w:type="lastRow"><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="6"/></w:tcBorders></w:tcPr></w:tblStylePr>
  <w:tblStylePr w:type="firstCol"><w:rPr><w:i/></w:rPr></w:tblStylePr>
</w:style>
<w:style w:type="table" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Banded"/>
  <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="FFF2CC"/></w:tcPr></w:tblStylePr></w:style>`;

describe('table style conditional formatting', () => {
  const rows = (n: number) => Array.from({ length: n }, (_, i) => row(cell(`r${i}a`) + cell(`r${i}b`))).join('');

  it('applies header-row, banded-row and first-column parts according to tblLook', async () => {
    const t = await importTable(tbl(rows(4), '<w:tblStyle w:val="Banded"/><w:tblLook w:firstRow="1" w:firstColumn="1" w:noHBand="0" w:noVBand="1"/>'), STYLE);
    const bg = t.rows.map((r) => r.cells[0].backgroundColor);
    // Row 0 is the header; banding starts after it (rows 1 and 3 are band 1, row 2 band 2).
    expect(bg).toEqual(['#336699', '#eeeeee', undefined, '#eeeeee']);
    expect(t.rows[0].cells[0].content[0]).toMatchObject({ text: 'r0a', bold: true, italic: true, color: '#ffffff' });
    expect(t.rows[0].cells[1].content[0]).toMatchObject({ bold: true, color: '#ffffff' });
    expect((t.rows[0].cells[1].content[0] as { italic?: boolean }).italic).toBeUndefined();
    expect(t.rows[1].cells[0].content[0]).toMatchObject({ italic: true });
    expect((t.rows[1].cells[1].content[0] as { italic?: boolean }).italic).toBeUndefined();
  });

  it('applies no conditional part that tblLook switches off, and no banding with noHBand', async () => {
    const t = await importTable(tbl(rows(3), '<w:tblStyle w:val="Banded"/><w:tblLook w:firstRow="0" w:firstColumn="0" w:noHBand="1" w:noVBand="1"/>'), STYLE);
    expect(t.rows.map((r) => r.cells[0].backgroundColor)).toEqual([undefined, undefined, undefined]);
    expect((t.rows[0].cells[0].content[0] as { bold?: boolean }).bold).toBeUndefined();
  });

  it('puts the style borders in the table, applies lastRow borders to that row and lets direct cell values win', async () => {
    const body = tbl(
      row(cell('a')) + row(cell('b', '<w:shd w:val="clear" w:fill="00FF00"/>')) + row(cell('c', '<w:tcBorders><w:top w:val="single" w:sz="24"/></w:tcBorders>')),
      '<w:tblStyle w:val="Banded"/><w:tblLook w:lastRow="1" w:firstRow="0" w:noHBand="0"/>'
    );
    const t = await importTable(body, STYLE);
    expect(t.borders).toMatchObject({ top: { style: 'single' }, insideH: { style: 'single', color: '#999999' } });
    expect(t.rows[1].cells[0].backgroundColor).toBe('#00ff00'); // direct shading beats band 2... and band 1
    expect(t.rows[2].cells[0].borders?.top).toEqual({ style: 'single', widthPx: 4 }); // direct tcBorders beat the style's lastRow
  });

  it('resolves a table style through its basedOn chain', async () => {
    const t = await importTable(tbl(rows(2), '<w:tblStyle w:val="Child"/><w:tblLook w:firstRow="1" w:noHBand="0"/>'), STYLE);
    expect(t.rows[0].cells[0].backgroundColor).toBe('#336699'); // inherited firstRow
    expect(t.rows[1].cells[0].backgroundColor).toBe('#fff2cc'); // overridden band1Horz
    expect(t.borders?.insideH?.color).toBe('#999999');
  });

  it('keeps a table style that is based on another, with the base', async () => {
    const result = await importDocx(await docx(tbl(rows(1), '<w:tblStyle w:val="Child"/>'), STYLE));
    expect(result.namedStyles?.Child.basedOn).toBe('Banded');
    expect(result.namedStyles?.Banded.kind).toBe('table');
  });

  it('does not let a non-default paragraph style be overridden by the table style for properties it leaves at the default', async () => {
    // docDefaults make text 11pt; the table style makes the header bold. Normal (default) must not undo it.
    const styles =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>' +
      STYLE;
    const t = await importTable(tbl(rows(1), '<w:tblStyle w:val="Banded"/><w:tblLook w:firstRow="1"/>'), styles);
    expect(t.rows[0].cells[0].content[0]).toMatchObject({ bold: true, fontSize: 11 });
  });

  it('keeps table styles in namedStyles with their conditional parts', async () => {
    const result = await importDocx(await docx(tbl(rows(1), '<w:tblStyle w:val="Banded"/>'), STYLE));
    expect(result.namedStyles?.Banded).toMatchObject({
      kind: 'table',
      table: {
        borders: { top: { style: 'single' } },
        conditions: { firstRow: { backgroundColor: '#336699', text: { bold: true, color: '#ffffff' } }, band1Horz: { backgroundColor: '#eeeeee' } },
      },
    });
    expect(result.namedStyles?.Banded.text).toBeUndefined();
  });
});

describe('applicableParts', () => {
  const style: DocTableStyle = {
    whole: { backgroundColor: '#ffffff' },
    rowBandSize: 2,
    conditions: { band1Horz: { backgroundColor: '#111111' }, band2Horz: { backgroundColor: '#222222' }, lastCol: { verticalAlign: 'bottom' }, neCell: { backgroundColor: '#333333' } },
  };
  const colorOf = (row: number, look = {}) => applicableParts(style, look, { row, rows: 6, col: 0, colspan: 1, cols: 2 }).map((p) => p.backgroundColor);
  it('bands in groups of rowBandSize', () => {
    expect([0, 1, 2, 3, 4].map((r) => colorOf(r).pop())).toEqual(['#111111', '#111111', '#222222', '#222222', '#111111']);
  });
  it('applies corner parts after row and column parts', () => {
    const parts = applicableParts(style, { firstRow: true, lastColumn: true, noHBand: true }, { row: 0, rows: 3, col: 1, colspan: 1, cols: 2 });
    expect(parts.map((p) => p.backgroundColor ?? p.verticalAlign)).toEqual(['#ffffff', 'bottom', '#333333']);
  });
});

describe('table formatting export', () => {
  const BASE: Omit<DocSection, 'blocks'> = { pageWidthPx: 816, pageHeightPx: 1056, marginTopPx: 96, marginRightPx: 96, marginBottomPx: 96, marginLeftPx: 96, orientation: 'portrait' };
  const text = (t: string) => [{ type: 'text' as const, text: t }];
  async function roundTrip(block: Table, namedStyles?: Record<string, never> | Record<string, import('@b32nio/spindle-shared/export/docx').DocNamedStyle>) {
    const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [block] }], ...(namedStyles && { namedStyles }) });
    const zip = await JSZip.loadAsync(exported.bytes);
    return { xml: await zip.file('word/document.xml')!.async('string'), styles: await zip.file('word/styles.xml')!.async('string'), imported: await importDocx(exported.bytes) };
  }

  const block: Table = {
    type: 'table',
    colWidthsPx: [96, 192],
    borders: { top: { style: 'double', widthPx: 2, color: '#ff0000' }, insideH: { style: 'dotted' } },
    rows: [
      {
        repeatHeader: true,
        cantSplit: true,
        cells: [
          { content: text('h1'), backgroundColor: '#ddeeff', verticalAlign: 'center', colwidthPx: [96], borders: { bottom: { style: 'single', widthPx: 3, color: '#00ff00' }, left: { style: 'none' } } },
          { content: text('h2'), colwidthPx: [192] },
        ],
      },
      { cells: [{ content: text('a'), colspan: 2, colwidthPx: [96, 192] }] },
    ],
  };

  it('writes borders, shading, header row, cantSplit, vAlign and widths', async () => {
    const { xml } = await roundTrip(block);
    expect(xml).toContain('<w:tblHeader/>');
    expect(xml).toContain('<w:cantSplit/>');
    expect(xml).toContain('<w:vAlign w:val="center"/>');
    expect(xml).toMatch(/<w:shd [^>]*w:fill="DDEEFF"/);
    expect(xml).toMatch(/<w:tblBorders><w:top w:val="double" w:color="FF0000" w:sz="12"\/>/);
    expect(xml).toMatch(/<w:tcW w:type="dxa" w:w="1440"\/>/);
    expect(xml).toMatch(/<w:tcW w:type="dxa" w:w="4320"\/>/);
  });

  it('survives an export and import unchanged', async () => {
    const { imported } = await roundTrip(block);
    const t = imported.sections[0].blocks[0] as Table;
    expect(t.colWidthsPx).toEqual([96, 192]);
    // Sides written as none read back as unlisted (a table side that is not listed has no line).
    expect(t.borders).toEqual({ top: { style: 'double', widthPx: 2, color: '#ff0000' }, insideH: { style: 'dotted' } });
    expect(t.rows[0]).toMatchObject({ repeatHeader: true, cantSplit: true });
    expect(t.rows[0].cells[0]).toMatchObject({ backgroundColor: '#ddeeff', verticalAlign: 'center', colwidthPx: [96], borders: { bottom: { style: 'single', widthPx: 3, color: '#00ff00' }, left: { style: 'none' } } });
    expect(t.rows[1].cells[0]).toMatchObject({ colspan: 2, colwidthPx: [96, 192] });
  });

  it('writes a table without borders with none on every side, so Word draws no grid', async () => {
    const { xml, imported } = await roundTrip({ type: 'table', borders: {}, rows: [{ cells: [{ content: text('a') }] }] });
    expect(xml).not.toContain('w:val="single"');
    expect((imported.sections[0].blocks[0] as Table).borders).toEqual({});
  });

  it('writes a table style with its look and conditional parts, and reads them back', async () => {
    const style: import('@b32nio/spindle-shared/export/docx').DocNamedStyle = {
      name: 'Banded',
      kind: 'table',
      table: {
        borders: { top: { style: 'single' } },
        rowBandSize: 2,
        whole: { text: { fontSize: 10 } },
        conditions: { firstRow: { backgroundColor: '#336699', verticalAlign: 'center', text: { bold: true, color: '#ffffff' } }, band1Horz: { borders: { bottom: { style: 'dashed', color: '#999999' } } } },
      },
    };
    const { styles, xml, imported } = await roundTrip({ ...block, namedStyleId: 'Banded', look: { firstRow: true, noVBand: true } }, { Banded: style });
    expect(xml).toContain('<w:tblStyle w:val="Banded"/>');
    expect(xml).toMatch(/<w:tblLook w:firstRow="true" w:lastRow="false" w:firstColumn="false" w:lastColumn="false" w:noHBand="false" w:noVBand="true"\/>/);
    expect(styles).toContain('w:tblStylePr w:type="firstRow"');
    expect(styles).toContain('w:tblStyleRowBandSize w:val="2"');
    expect(imported.namedStyles?.Banded).toEqual(style);
    expect((imported.sections[0].blocks[0] as Table).namedStyleId).toBe('Banded');
  });

  it('drops a reference to a table style the document does not define', async () => {
    const { xml } = await roundTrip({ ...block, namedStyleId: 'Gone' });
    expect(xml).not.toContain('w:tblStyle');
  });
});
