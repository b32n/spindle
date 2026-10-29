import JSZip from 'jszip';
import { countMatches, createLossReport, type ImportLossItem, type ImportLossReport } from '@b32nio/spindle-shared';

const MAX_WHERE = 5;

/** Sheet name + part path pairs from workbook.xml and its relationships. */
async function sheetParts(zip: JSZip): Promise<Array<{ name: string; path: string; hidden: boolean }>> {
  const wb = await zip.file('xl/workbook.xml')?.async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!wb) return [];
  const targets = new Map<string, string>();
  for (const m of (rels ?? '').matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]*)"/.exec(m[0])?.[1];
    const target = /\bTarget="([^"]*)"/.exec(m[0])?.[1];
    if (id && target) targets.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`);
  }
  const out: Array<{ name: string; path: string; hidden: boolean }> = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /\bname="([^"]*)"/.exec(m[0])?.[1] ?? '?';
    const rid = /\br:id="([^"]*)"/.exec(m[0])?.[1];
    const path = rid ? targets.get(rid) : undefined;
    const state = /\bstate="([^"]*)"/.exec(m[0])?.[1];
    if (path) out.push({ name: unescapeXml(name), path, hidden: state === 'hidden' || state === 'veryHidden' });
  }
  return out;
}

function unescapeXml(s: string): string {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

/**
 * Scan an .xlsx for features the importer does not bring across. Cheap (zip +
 * pattern counts, no workbook model), never throws on a malformed file — an
 * unreadable zip yields an empty report and the import itself reports the error.
 */
export async function detectXlsxLoss(bytes: Uint8Array): Promise<ImportLossReport> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return createLossReport('xlsx', []);
  }
  const names = Object.keys(zip.files);
  const items: ImportLossItem[] = [];
  const add = (id: string, label: string, count: number, effect: string, where?: string[]) => {
    if (count > 0) items.push({ id, label, count, effect, ...(where && where.length ? { where: where.slice(0, MAX_WHERE) } : {}) });
  };

  // ---- per-sheet XML ------------------------------------------------------
  const perSheet = {
    cf: [] as Array<[string, number]>,
    dv: [] as Array<[string, number]>,
    filter: [] as Array<[string, number]>,
    protect: [] as Array<[string, number]>,
    outline: [] as Array<[string, number]>,
    spark: [] as Array<[string, number]>,
  };
  const parts = await sheetParts(zip);
  for (const part of parts) {
    const xml = await zip.file(part.path)?.async('string');
    if (!xml) continue;
    perSheet.cf.push([part.name, countMatches(xml, /<conditionalFormatting[\s>/]/g)]);
    perSheet.dv.push([part.name, countMatches(xml, /<dataValidation[\s>/]/g)]);
    perSheet.filter.push([part.name, countMatches(xml, /<autoFilter[\s>/]/g)]);
    perSheet.protect.push([part.name, countMatches(xml, /<sheetProtection\b[^>]*\bsheet="(?:1|true)"/g)]);
    perSheet.outline.push([part.name, countMatches(xml, /\boutlineLevel(?:Row|Col)?="[1-9]/g)]);
    perSheet.spark.push([part.name, countMatches(xml, /<(?:x14:)?sparklineGroup[\s>/]/g)]);
  }
  const total = (rows: Array<[string, number]>) => rows.reduce((n, [, c]) => n + c, 0);
  const where = (rows: Array<[string, number]>) => rows.filter(([, c]) => c > 0).map(([n]) => n);
  add('xlsx.conditionalFormatting', 'conditional formatting rule(s)', total(perSheet.cf), 'not imported', where(perSheet.cf));
  add('xlsx.dataValidation', 'data validation rule(s) (dropdowns, limits)', total(perSheet.dv), 'not imported', where(perSheet.dv));
  add('xlsx.autoFilter', 'filter(s)', total(perSheet.filter), 'not imported', where(perSheet.filter));
  add('xlsx.sheetProtection', 'protected sheet(s)', total(perSheet.protect), 'protection is dropped', where(perSheet.protect));
  add('xlsx.outline', 'grouped (outlined) row/column(s)', total(perSheet.outline), 'groups are flattened', where(perSheet.outline));
  add('xlsx.sparklines', 'sparkline group(s)', total(perSheet.spark), 'not imported', where(perSheet.spark));
  add('xlsx.hiddenSheets', 'hidden sheet(s)', parts.filter((p) => p.hidden).length, 'become visible', parts.filter((p) => p.hidden).map((p) => p.name));

  // ---- workbook-level -----------------------------------------------------
  const wbXml = (await zip.file('xl/workbook.xml')?.async('string')) ?? '';
  const definedNames = [...wbXml.matchAll(/<definedName\b[^>]*\bname="([^"]*)"/g)].map((m) => m[1]).filter((n) => !n.startsWith('_xlnm.'));
  add('xlsx.definedNames', 'named range(s)/name(s)', definedNames.length, 'not imported; formulas using them break', definedNames);
  add('xlsx.workbookProtection', 'workbook protection', countMatches(wbXml, /<workbookProtection\b[^>]*\b(?:lockStructure|lockWindows)="(?:1|true)"/g), 'protection is dropped');

  // ---- parts that exist only when the feature is used ---------------------
  const count = (re: RegExp) => names.filter((n) => re.test(n)).length;
  add('xlsx.charts', 'chart(s)', count(/^xl\/charts\/chart\d*\.xml$/), 'not imported');
  add('xlsx.images', 'image(s)/drawing(s)', count(/^xl\/media\//), 'not imported');
  add('xlsx.comments', 'sheet(s) with comments/notes', count(/^xl\/(comments\d*\.xml|threadedComments\/)/), 'comments are dropped');
  add('xlsx.tables', 'table(s)', count(/^xl\/tables\/table\d*\.xml$/), 'imported as plain cells (no table styling or structured references)');
  add('xlsx.pivotTables', 'pivot table(s)', count(/^xl\/pivotTables\/pivotTable\d*\.xml$/), 'only the last computed values are kept');
  add('xlsx.macros', 'macro project', count(/^xl\/vbaProject\.bin$/), 'macros are dropped');
  add('xlsx.externalLinks', 'external workbook link(s)', count(/^xl\/externalLinks\/externalLink\d*\.xml$/), 'not imported');
  add('xlsx.slicers', 'slicer(s)', count(/^xl\/slicers\//), 'not imported');

  // ---- rich text inside cells --------------------------------------------
  const sst = (await zip.file('xl/sharedStrings.xml')?.async('string')) ?? '';
  add('xlsx.richText', 'cell(s) with mixed formatting inside the text', countMatches(sst, /<si>\s*<r>\s*<rPr>/g), 'imported as plain text');

  return createLossReport('xlsx', items);
}
