// A tiny .xlsx builder for tests: the sheet data of a workbook this editor exported is replaced by hand-written rows, so a case
// is what Excel writes (`_xlfn.` prefixes, `t="array"` with `cm`) with the results Excel saved.
import JSZip from 'jszip';
import { exportToXlsx } from './xlsx';
import { WorkbookImpl } from '../workbook';

const XML_ESC: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
const esc = (s: string): string => s.replace(/[&<>"]/g, (c) => XML_ESC[c]);

/** `<c>` for a formula: `f` is the `<f>` element's inner text, `attrs` its attributes, `value` the saved result (numbers and strings). */
export function fcell(ref: string, f: string, value: number | string | null, attrs = '', cellAttrs = ''): string {
  const t = typeof value === 'string' ? ' t="str"' : '';
  const v = value === null ? '' : `<v>${esc(String(value))}</v>`;
  return `<c r="${ref}"${t}${cellAttrs}><f${attrs ? ` ${attrs}` : ''}>${esc(f)}</f>${v}</c>`;
}
export const ncell = (ref: string, value: number): string => `<c r="${ref}"><v>${value}</v></c>`;

/** An .xlsx whose first sheet holds `cellsByRow` (row number -> cell XML), made from an export of this editor. */
export async function xlsxWith(cellsByRow: Record<number, string[]>): Promise<Uint8Array> {
  const base = new WorkbookImpl('wb', 'WB');
  base.setCellValue(undefined, 0, 0, 1);
  const zip = await JSZip.loadAsync(await exportToXlsx(base));
  const rows = Object.entries(cellsByRow)
    .map(([r, cells]) => `<row r="${r}">${cells.join('')}</row>`)
    .join('');
  const sheet = (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).replace(/<sheetData>[\s\S]*<\/sheetData>/, `<sheetData>${rows}</sheetData>`);
  zip.file('xl/worksheets/sheet1.xml', sheet);
  return zip.generateAsync({ type: 'uint8array' });
}
