import JSZip from 'jszip';
import type { ImportLossReport } from '@b32nio/spindle-shared/export/xlsx';
import { exportToXlsx, importFromXlsx, inspectXlsx } from './xlsx';
import { WorkbookImpl } from '../workbook';
import { getRegisteredFunctionNames } from '../formula-parser/functions';

const ids = (r: ImportLossReport) => r.items.map((i) => i.id).sort();
const item = (r: ImportLossReport, id: string) => r.items.find((i) => i.id === id);

async function patched(
  wb: WorkbookImpl,
  edit: (zip: JSZip) => Promise<void> | void
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await exportToXlsx(wb));
  await edit(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

async function patchSheet(wb: WorkbookImpl, fn: (xml: string) => string): Promise<Uint8Array> {
  return patched(wb, async (zip) => {
    zip.file(
      'xl/worksheets/sheet1.xml',
      fn(await zip.file('xl/worksheets/sheet1.xml')!.async('string'))
    );
  });
}

describe('inspectXlsx', () => {
  it('reports nothing for a workbook this editor exported (values, formulas, formats, styles)', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 'text');
    wb.setCellValue(undefined, 0, 1, 1234.5);
    wb.setFormula(undefined, 0, 2, '=SUM(B1,2)+IF(TRUE(),1,2)');
    const formats = [
      { type: 'number', decimalPlaces: 2, useThousandsSeparator: true },
      { type: 'currency', currencyCode: 'USD', decimalPlaces: 2, useThousandsSeparator: true },
      { type: 'percentage', decimalPlaces: 1 },
      { type: 'date', dateFormat: 'YYYY-MM-DD' },
      { type: 'datetime', dateFormat: 'YYYY-MM-DD', timeFormat: 'HH:mm:ss' },
    ] as const;
    formats.forEach((format, i) => {
      wb.setCell(undefined, 1, i, {
        value: i,
        styleId: wb
          .getStylePool()
          .getOrCreate({ bold: true, fontColor: '#ff0000', backgroundColor: '#ffff00' }),
        formatId: wb.getFormatPool().getOrCreate(format),
      });
    });
    const bytes = await exportToXlsx(wb);
    expect(await inspectXlsx(bytes)).toEqual({ format: 'xlsx', items: [] });
    expect(await inspectXlsx(bytes, { detail: 'full' })).toEqual({ format: 'xlsx', items: [] });
  });

  it('is summary by default and full on request', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 1);
    const bytes = await patchSheet(wb, (xml) =>
      xml
        .replace('<dimension', '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr><dimension')
        .replace('</sheetData>', '</sheetData><autoFilter ref="A1:A2"/>')
    );
    expect(ids(await inspectXlsx(bytes))).toEqual(['xlsx.autoFilter']);
    expect(ids(await inspectXlsx(bytes, { detail: 'summary' }))).toEqual(['xlsx.autoFilter']);
    expect(ids(await inspectXlsx(bytes, { detail: 'full' }))).toEqual([
      'xlsx.autoFilter',
      'xlsx.sheet.tabColor',
    ]);
  });

  it('judges formula functions by the engine catalog, ignoring the future-function prefixes', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 1);
    const cells =
      '<row r="1"><c r="A1"><f>_xlfn.IFS(1,2)+SUM(1)</f><v>1</v></c><c r="B1"><f>REGEXEXTRACT("a","a")*REGEXREPLACE("a","a","b")</f><v>1</v></c><c r="C1"><f>_xlfn.NOSUCH(1,2,3)</f><v>1</v></c></row>';
    const bytes = await patchSheet(wb, (xml) =>
      xml.replace(/<sheetData>.*<\/sheetData>/s, `<sheetData>${cells}</sheetData>`)
    );
    const report = await inspectXlsx(bytes);
    expect(item(report, 'xlsx.formulas.unsupported')).toMatchObject({
      count: 2,
      where: ['REGEXEXTRACT', 'REGEXREPLACE', 'NOSUCH'],
    });
    for (const name of ['IFS', 'SUM', 'IF'])
      expect(getRegisteredFunctionNames().has(name)).toBe(true);
    for (const name of ['NOSUCH', 'REGEXEXTRACT', 'REGEXREPLACE'])
      expect(getRegisteredFunctionNames().has(name)).toBe(false);
  });

  it('reports structured references in the summary', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 1);
    const bytes = await patchSheet(wb, (xml) =>
      xml.replace(
        /<sheetData>.*<\/sheetData>/s,
        '<sheetData><row r="1"><c r="A1"><f>SUM(Table1[Amount])</f><v>1</v></c></row></sheetData>'
      )
    );
    expect(item(await inspectXlsx(bytes), 'xlsx.formulas.structuredRefs')).toMatchObject({
      count: 1,
    });
  });

  describe('number formats', () => {
    async function withCode(code: string | undefined): Promise<ImportLossReport> {
      const wb = new WorkbookImpl('wb', 'WB');
      wb.setCellValue(undefined, 0, 0, 1);
      const bytes = await patched(wb, async (zip) => {
        const styles = await zip.file('xl/styles.xml')!.async('string');
        const fmts =
          code === undefined
            ? ''
            : `<numFmts count="1"><numFmt numFmtId="164" formatCode="${code.replace(/"/g, '&quot;')}"/></numFmts>`;
        const xfs = `<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="${code === undefined ? 14 : 164}"/></cellXfs>`;
        zip.file(
          'xl/styles.xml',
          styles
            .replace('<fonts', () => `${fmts}<fonts`)
            .replace(/<cellXfs[\s\S]*?<\/cellXfs>/, () => xfs)
        );
        const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
        zip.file('xl/worksheets/sheet1.xml', sheet.replace('<c r="A1"', '<c r="A1" s="1"'));
      });
      return inspectXlsx(bytes);
    }

    it.each([
      '0.00', '#,##0.00', '0%', '0.0%', 'yyyy-mm-dd', 'mm/dd/yyyy', 'hh:mm:ss', '"$"#,##0.00', '0.00E+00', 'General',
      '0.0,,"M"', '@', '[h]:mm:ss', '# ?/?', 'd-mmm-yy', '[>100]"big";"small"', '0.0 "kg"', '000-00', '[$€-407] #,##0.00', '[$-409]mmmm d, yyyy',
      '[$-F800]dddd, mmmm dd, yyyy', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    ])(
      'accepts %s: it survives an import and an export unchanged and the engine shows it like Excel',
      async (code) => {
        expect(ids(await withCode(code))).toEqual([]);
      }
    );

    it.each(['[DBNum1]0', '[$-411]ge.m.d', '[$-407]dddd, d. mmmm yyyy', '0 abc', 'yyyy 0', '0;0;0;0;0'])(
      'reports %s as approximated: the code is kept, but the engine cannot show it like Excel',
      async (code) => {
        expect(item(await withCode(code), 'xlsx.numberFormats.approximated')).toMatchObject({ count: 1, where: [code] });
      }
    );

    it('accepts a built-in date id: its en-US code is kept as a custom format', async () => {
      expect(ids(await withCode(undefined))).toEqual([]);
    });

    it('still reports a built-in id that has no code (27: an era date of the Asian locales)', async () => {
      const wb = new WorkbookImpl('wb', 'WB');
      wb.setCellValue(undefined, 0, 0, 1);
      const bytes = await patched(wb, async (zip) => {
        const styles = await zip.file('xl/styles.xml')!.async('string');
        zip.file('xl/styles.xml', styles.replace(/<cellXfs[\s\S]*?<\/cellXfs>/, () => '<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="27"/></cellXfs>'));
        const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
        zip.file('xl/worksheets/sheet1.xml', sheet.replace('<c r="A1"', '<c r="A1" s="1"'));
      });
      expect(item(await inspectXlsx(bytes), 'xlsx.numberFormats.approximated')).toMatchObject({ count: 1, where: ['built-in format 27'] });
    });
  });

  it('is what importFromXlsx returns as its loss report, at the summary level', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    wb.setCellValue(undefined, 0, 0, 1);
    const bytes = await patchSheet(wb, (xml) =>
      xml
        .replace('<dimension', '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr><dimension')
        .replace(
          /<sheetData>.*<\/sheetData>/s,
          '<sheetData><row r="1"><c r="A1"><f>SUM(Table1[a])+REGEXEXTRACT(1,2)</f><v>1</v></c></row></sheetData>'
        )
    );
    const target = new WorkbookImpl('t', 'T');
    const { loss } = await importFromXlsx(bytes, target);
    expect(loss).toEqual(await inspectXlsx(bytes));
    expect(ids(loss)).toEqual(['xlsx.formulas.structuredRefs', 'xlsx.formulas.unsupported']);
  });
});
