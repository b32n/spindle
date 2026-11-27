import JSZip from 'jszip';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { fromExcelNumberFormatCode, toExcelNumberFormatCode } from './xlsx-number-format';
import { compileFormat, formatDisplay, isValidFormatCode } from '../number-format';
import { WorkbookImpl } from '../workbook';
import type { CellFormat } from '../types';

// Number format codes that carry a backslash escape, through the path the editor uses (S2b: the reader keeps the backslash where
// it is what makes a character a literal, and drops it before a character Excel shows without quotes). Each code is written by
// ExcelJS as a file would have it, opened with `importFromXlsx`, shown by the S3 engine, written by `exportToXlsx` and read again.
//
//   raw        the code as the file writes it
//   editor     the code the editor holds (the importer's answer)
//   value      a number to show
//   shown      what Excel shows for the value
const CASES: Array<{ raw: string; editor: string; value: number; shown: string }> = [
  { raw: '\\$#,##0.00', editor: '$#,##0.00', value: 1234.5, shown: '$1,234.50' },
  { raw: '_(\\$* #,##0_)', editor: '_($* #,##0_)', value: 1234, shown: ' $1,234 ' },
  { raw: '0.0\\"', editor: '0.0\\"', value: 5, shown: '5.0"' },
  { raw: '0\\ \\m', editor: '0 \\m', value: 5, shown: '5 m' },
  { raw: '0\\ \\d', editor: '0 \\d', value: 7, shown: '7 d' },
  { raw: '\\*0\\*', editor: '\\*0\\*', value: 5, shown: '*5*' },
  { raw: '0\\.00', editor: '0\\.00', value: 5, shown: '0.05' },
  { raw: '0\\\\', editor: '0\\\\', value: 5, shown: '5\\' },
  { raw: '"C:\\Temp"0', editor: '"C:\\Temp"0', value: 5, shown: 'C:\\Temp5' },
  { raw: '0\\ \\k\\g', editor: '0 \\k\\g', value: 5, shown: '5 kg' },
  { raw: '0\\,0', editor: '0\\,0', value: 5, shown: '0,5' },
  { raw: '0\\%', editor: '0\\%', value: 5, shown: '5%' },
];

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const RELS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const attr = (code: string): string => code.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

/** A workbook whose A1..An hold the values of CASES, formatted with `codes` written as given (ExcelJS would rewrite a code on the way out, so the package is built by hand). */
async function fileWith(codes: string[]): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  const rels = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${RELS}/officeDocument" Target="xl/workbook.xml"/>`));
  zip.file('xl/workbook.xml', `<workbook ${NS} xmlns:r="${RELS}"><sheets><sheet name="F" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', rels(`<Relationship Id="rId1" Type="${RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${RELS}/styles" Target="styles.xml"/>`));
  zip.file(
    'xl/styles.xml',
    `<styleSheet ${NS}><numFmts count="${codes.length}">${codes.map((c, i) => `<numFmt numFmtId="${170 + i}" formatCode="${attr(c)}"/>`).join('')}</numFmts>` +
      `<cellXfs><xf/>${codes.map((_, i) => `<xf numFmtId="${170 + i}" applyNumberFormat="1"/>`).join('')}</cellXfs></styleSheet>`
  );
  zip.file('xl/worksheets/sheet1.xml', `<worksheet ${NS}><sheetData>${codes.map((_, i) => `<row r="${i + 1}"><c r="A${i + 1}" s="${i + 1}"><v>${CASES[i].value}</v></c></row>`).join('')}</sheetData></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

describe('number formats with backslash escapes, through the editor', () => {
  const formats = async (bytes: Uint8Array) => {
    const workbook = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(bytes, workbook);
    const pool = workbook.getFormatPool();
    return { workbook, formats: CASES.map((_, i) => pool.get(workbook.getCell(undefined, i, 0)?.formatId as string)) };
  };

  it('imports each code as the code the editor holds (custom, verbatim) and the engine reads it', async () => {
    const { formats: got } = await formats(await fileWith(CASES.map((c) => c.raw)));
    CASES.forEach((c, i) => {
      expect(got[i]).toEqual({ type: 'custom', pattern: c.editor });
      expect(isValidFormatCode(c.editor)).toBe(true);
      expect(isValidFormatCode(c.raw)).toBe(true);
    });
  });

  it('shows each value as Excel does, and the same as the code the file wrote', () => {
    for (const c of CASES) {
      expect(formatDisplay(c.value, c.editor)).toBe(c.shown);
      expect(formatDisplay(c.value, c.raw)).toBe(c.shown);
      expect(compileFormat(c.editor).sections.map((s) => s.kind)).toEqual(compileFormat(c.raw).sections.map((s) => s.kind));
    }
  });

  it('keeps a date code of escaped separators a date (a structured type), and writes the same code back', () => {
    const date = fromExcelNumberFormatCode('yyyy-mm-dd');
    expect(date).toEqual({ type: 'date', dateFormat: 'YYYY-MM-DD' });
    expect(toExcelNumberFormatCode(date)).toBe('yyyy-mm-dd');
    expect(formatDisplay(46024, 'yyyy\\-mm\\-dd')).toBe(formatDisplay(46024, 'yyyy-mm-dd'));
  });

  it('writes each code back as a valid code of the file, and reads it again the same', async () => {
    const { workbook } = await formats(await fileWith(CASES.map((c) => c.raw)));
    const out = await exportToXlsx(workbook);
    const zip = await JSZip.loadAsync(out);
    const styles = await (zip.file('xl/styles.xml') as JSZip.JSZipObject).async('string');
    // the attribute value is escaped as XML (a quote is `&quot;`), and what it holds is the editor's code
    const written = [...styles.matchAll(/<numFmt numFmtId="\d+" formatCode="([^"]*)"/g)].map((m) => m[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&'));
    for (const c of CASES) expect(written).toContain(c.editor);
    for (const code of written) expect(isValidFormatCode(code)).toBe(true);
    // (not through ExcelJS's reader: it drops every escape, which is what the S2b reader no longer does) a second import gives the same formats and the same values to show
    const second = new WorkbookImpl('wb2', 'WB2');
    await importFromXlsx(out, second);
    const pool = second.getFormatPool();
    CASES.forEach((c, i) => {
      const format = pool.get(second.getCell(undefined, i, 0)?.formatId as string) as CellFormat;
      expect(format).toEqual({ type: 'custom', pattern: c.editor });
      expect(formatDisplay(c.value, toExcelNumberFormatCode(format))).toBe(c.shown);
    });
  });
});
