import JSZip from 'jszip';
import type { SpreadsheetCellStyle } from '@b32nio/spindle-shared/export/xlsx';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { MAX_IMPORTED_STYLES, StyleImport, fromSpreadsheetStyle, toDefaultFormat, toSpreadsheetStyle } from './xlsx-style';
import { StylePool } from '../style-pool';
import { FormatPool } from '../format-pool';
import { WorkbookImpl } from '../workbook';
import type { BorderKind, CellStyle } from '../types';
import { APPROXIMATED_BORDER_KINDS, CSS_OF_KIND, inferBorderKind, isBorderKind, keptBorderKind, resolveBorderKind } from '../border-kinds';

const KINDS = Object.keys(CSS_OF_KIND) as BorderKind[];

describe('border kinds', () => {
  it('has thirteen kinds, each recognised', () => {
    expect(KINDS).toHaveLength(13);
    for (const k of KINDS) expect(isBorderKind(k)).toBe(true);
    for (const bad of ['bogus', 'toString', '__proto__', 3, undefined, null]) expect(isBorderKind(bad)).toBe(false);
  });

  it('infers what the editor tools make and keeps a kind only while the shorthand still matches it', () => {
    expect(inferBorderKind(1, 'solid')).toBe('thin');
    expect(inferBorderKind(2, 'solid')).toBe('medium');
    expect(inferBorderKind(3, 'solid')).toBe('thick');
    expect(inferBorderKind(1, 'dashed')).toBe('dashed');
    expect(inferBorderKind(1, 'dotted')).toBe('dotted');
    expect(inferBorderKind(3, 'double')).toBe('double');
    expect(resolveBorderKind(1, 'dashed', 'dashDot')).toBe('dashDot');
    expect(resolveBorderKind(2, 'solid', 'dashDot')).toBe('medium'); // the shorthand was edited since
    expect(resolveBorderKind(1, 'dashed', undefined)).toBe('dashed');
    expect(resolveBorderKind(1, 'dashed', 'constructor' as BorderKind)).toBe('dashed');
  });

  it('keeps a kind next to the shorthand only where the shorthand alone cannot give it back', () => {
    const kept = KINDS.filter((k) => keptBorderKind(k) !== undefined);
    expect(kept.sort()).toEqual(['dashDot', 'dashDotDot', 'hair', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot']);
    expect([...APPROXIMATED_BORDER_KINDS].sort()).toEqual(['hair', 'slantDashDot']);
  });
});

describe('styles between the file and the pool', () => {
  const full: SpreadsheetCellStyle = {
    bold: true, italic: true, underline: true, underlineStyle: 'doubleAccounting', strikethrough: true,
    fontFamily: 'Arial', fontSize: 12, fontColor: '#112233', backgroundColor: '#445566', fillPattern: 'darkGrid', patternColor: '#778899',
    horizontalAlign: 'right', verticalAlign: 'top', wrapText: true, indent: 2, shrinkToFit: true, textRotation: 30,
  };

  it('round-trips every kind on every side with its colour', () => {
    for (const kind of KINDS) {
      const style: SpreadsheetCellStyle = {
        border: { top: { style: kind, color: '#ff0000' }, right: { style: kind, color: '#00ff00' }, bottom: { style: kind }, left: { style: kind, color: '#0000ff' } },
      };
      const pooled = fromSpreadsheetStyle(style);
      expect(pooled.borderTop).toBe(`${CSS_OF_KIND[kind].width}px ${CSS_OF_KIND[kind].style} #ff0000`);
      expect(pooled.borderBottom).toContain('#000000');
      expect(toSpreadsheetStyle(pooled)).toEqual({
        border: { top: { style: kind, color: '#ff0000' }, right: { style: kind, color: '#00ff00' }, bottom: { style: kind, color: '#000000' }, left: { style: kind, color: '#0000ff' } },
      });
    }
  });

  it('round-trips fonts, pattern fills, indent, shrink-to-fit, underline variants next to strike', () => {
    const pooled = fromSpreadsheetStyle(full);
    expect(pooled).toMatchObject({ textDecoration: 'underline', underlineKind: 'doubleAccounting', strikethrough: true, fillPattern: 'darkGrid', patternColor: '#778899', backgroundColor: '#445566', indent: 2, shrinkToFit: true });
    expect(toSpreadsheetStyle(pooled)).toEqual(full);
  });

  it('keeps strike alone as line-through, and ignores a pattern without a colour', () => {
    expect(fromSpreadsheetStyle({ strikethrough: true })).toEqual({ textDecoration: 'line-through' });
    expect(toSpreadsheetStyle({ textDecoration: 'line-through' })).toEqual({ strikethrough: true });
    expect(toSpreadsheetStyle({ strikethrough: true })).toEqual({ strikethrough: true });
    expect(fromSpreadsheetStyle({ fillPattern: 'darkGrid' })).toEqual({});
    expect(toSpreadsheetStyle({ fillPattern: 'darkGrid' })).toEqual({});
  });

  it('writes a border with a shorthand it cannot read as thin, and a named colour as none', () => {
    expect(toSpreadsheetStyle({ borderTop: 'odd' }).border?.top).toEqual({ style: 'thin' });
    expect(toSpreadsheetStyle({ borderTop: '1px solid red' }).border?.top).toEqual({ style: 'thin', color: undefined });
  });
});

describe('StyleImport', () => {
  it('converts a style once per object and shares the id', () => {
    const imp = new StyleImport(new StylePool(), new FormatPool());
    const a = { bold: true };
    expect(imp.style(a)).toBe(imp.style(a));
    expect(imp.style({ bold: true })).toBe(imp.style(a));
    const f = { code: '0.00' };
    expect(imp.format(f)).toBe(imp.format(f));
    expect(imp.format({ code: 'General' })).toBeUndefined();
  });

  it('stops at the limit, keeps the styles already pooled, and counts the others', () => {
    const imp = new StyleImport(new StylePool(), new FormatPool(), 2);
    expect(imp.style({ fontSize: 1 })).toBeDefined();
    expect(imp.style({ fontSize: 2 })).toBeDefined();
    expect(imp.style({ fontSize: 3 })).toBeUndefined();
    expect(imp.style({ fontSize: 2 })).toBeDefined(); // already there
    expect(imp.format({ code: '0.0' })).toBeDefined();
    expect(imp.format({ code: '0.00' })).toBeDefined();
    expect(imp.format({ code: '0.000' })).toBeUndefined();
    expect(imp.format({ code: '0.00' })).toBeDefined();
    expect(imp.cut).toEqual({ styles: 1, formats: 1, lines: 0 });
    expect(MAX_IMPORTED_STYLES).toBe(65_536);
  });

  it('makes a column format from its style and format, and nothing from nothing', () => {
    const imp = new StyleImport(new StylePool(), new FormatPool());
    expect(imp.line({})).toBeUndefined();
    expect(imp.line({ style: { bold: true } })).toEqual({ styleId: expect.any(String) });
    expect(imp.line({ numberFormat: { code: '0.00' } })).toEqual({ formatId: expect.any(String) });
    expect(imp.line({ style: { bold: true }, numberFormat: { code: '0.00' } })).toEqual({ styleId: expect.any(String), formatId: expect.any(String) });
  });

  it('exports a line format only when what it names exists', () => {
    const styles = new StylePool();
    const formats = new FormatPool();
    const s = styles.getOrCreate({ bold: true });
    const f = formats.getOrCreate({ type: 'number', decimalPlaces: 2 });
    expect(toDefaultFormat({ styleId: s, formatId: f }, styles, formats)).toEqual({ style: { bold: true }, numberFormat: { code: expect.any(String) } });
    expect(toDefaultFormat({ styleId: s }, styles, formats)).toEqual({ style: { bold: true } });
    expect(toDefaultFormat({ formatId: f }, styles, formats)).toEqual({ numberFormat: { code: expect.any(String) } });
    expect(toDefaultFormat({ styleId: 'gone', formatId: 'gone' }, styles, formats)).toBeUndefined();
  });
});

async function roundTrip(wb: WorkbookImpl): Promise<WorkbookImpl> {
  const out = new WorkbookImpl('rt', 'RT');
  await importFromXlsx(await exportToXlsx(wb), out);
  return out;
}

describe('column and row default formats through the file', () => {
  const bold: CellStyle = { bold: true, backgroundColor: '#ffeecc', borderBottom: '1px dashed #123456', borderBottomKind: 'dashDot', fillPattern: 'lightGrid', patternColor: '#000000' };

  it('writes and reads them, with the number format, and a new cell in the column starts with them', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 'x');
    const sheet = wb.getSheet(undefined)!;
    const style = wb.getStylePool().getOrCreate(bold);
    const format = wb.getFormatPool().getOrCreate({ type: 'percentage', decimalPlaces: 1 });
    sheet.setColFormat(2, { styleId: style, formatId: format });
    sheet.setRowFormat(4, { styleId: style });
    const back = await roundTrip(wb);
    const s2 = back.getSheet(undefined)!;
    const col = s2.getColFormat(2)!;
    expect(back.getStylePool().get(col.styleId!)).toMatchObject(bold);
    expect(back.getFormatPool().get(col.formatId!)).toMatchObject({ type: 'percentage' });
    expect(back.getStylePool().get(s2.getRowFormat(4)!.styleId!)).toMatchObject(bold);
    expect(s2.getColFormat(1)).toBeUndefined();
    back.setCellValue(undefined, 8, 2, 5);
    expect(back.getCell(undefined, 8, 2)?.styleId).toBe(col.styleId);
  });

  it('writes a default width and height that are not the editor\'s, and none that are', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 'x');
    const sheet = wb.getSheet(undefined)!;
    expect((await roundTrip(wb)).getSheet(undefined)!.config.defaultColWidth).toBeUndefined();
    sheet.config.defaultColWidth = 64;
    sheet.config.defaultRowHeight = 30;
    const back = (await roundTrip(wb)).getSheet(undefined)!;
    expect(back.config.defaultColWidth).toBe(64);
    expect(back.config.defaultRowHeight).toBe(30);
    expect(back.getColWidth(50)).toBe(64);
  });

  it('keeps default widths stable through the file for every width', async () => {
    for (let px = 20; px <= 400; px += 7) {
      const wb = new WorkbookImpl('wb', 'WB');
      wb.setCellValue(undefined, 0, 0, 'x');
      wb.getSheet(undefined)!.config.defaultColWidth = px;
      const once = (await roundTrip(wb)).getSheet(undefined)!.config.defaultColWidth;
      const twice = (await roundTrip(await roundTrip(wb))).getSheet(undefined)!.config.defaultColWidth;
      expect(Math.abs((once ?? 100) - px)).toBeLessThanOrEqual(1);
      expect(twice).toBe(once);
    }
  });

  it('drops formats of columns and rows past the sheet and reports them; a full-width column format costs no more than the sheet', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 'x');
    const bytes = await exportToXlsx(wb);
    const zip = await JSZip.loadAsync(bytes);
    const styles = await zip.file('xl/styles.xml')!.async('string');
    zip.file('xl/styles.xml', styles.replace(/<cellXfs count="\d+">/, '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>').replace(/<fills count="\d+">/, '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>').replace(/<xf [^>]*\/>(?=<\/cellXfs>)/, '').replace(/<\/cellXfs>[\s\S]*?<\/cellXfs>/, '</cellXfs>'));
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const rows = Array.from({ length: 1500 }, (_, i) => `<row r="${i + 1}" s="1" customFormat="1"/>`).join('');
    zip.file('xl/worksheets/sheet1.xml', sheet.replace('<sheetData>', '<cols><col min="1" max="16384" width="9" style="1" customWidth="1"/></cols><sheetData>').replace(/<sheetData>[\s\S]*<\/sheetData>/, `<sheetData>${rows}</sheetData>`).replace(/<cols>.*?<\/cols>(?=.*<cols>)/, ''));
    const out = new WorkbookImpl('o', 'O');
    const { loss } = await importFromXlsx(await zip.generateAsync({ type: 'uint8array' }), out);
    const s = out.getSheet(undefined)!;
    // the 100 default columns and the 1000 default rows (the sheet's extent) keep the format; the rest is reported
    expect(s.config.colStyles?.size).toBeLessThanOrEqual(16384);
    expect(s.getColFormat(0)).toBeDefined();
    expect(s.getRowFormat(999)).toBeDefined();
    expect(s.getRowFormat(1000)).toBeUndefined();
    expect(loss.items.some((i) => i.id === 'xlsx.package.sizeLimit' && i.label.includes('past the last used row or column'))).toBe(true);
  });
});
