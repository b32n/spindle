import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { parseXml, type XmlElement } from '../xml';

// Helpers for the reader tests: packages built in memory and the corpus files, so each test states only what it is about.

const CORPUS = resolve(__dirname, '../../../../../fixtures/xlsx');

/** A corpus file's bytes. */
export const corpus = (filename: string): Uint8Array => new Uint8Array(readFileSync(resolve(CORPUS, filename)));

/**
 * The table of number format codes the reader is checked against, as data (`fixtures/xlsx/data/number-format-cases.json`, read by this
 * package's cell-style test and by the parity test in sheets-core, so neither imports the other's source): `[raw, editor, isDate, why]`.
 * `raw` is the code as the file writes it, `editor` the code the editor gets (`unescapeLiterals`: the backslash goes before a character
 * Excel shows without quotes and stays before everything else, and nothing in quotes is touched), `isDate` whether it shows dates.
 */
export const numberFormatCases = (): Array<[raw: string, editor: string, isDate: boolean, why: string]> =>
  JSON.parse(readFileSync(resolve(CORPUS, 'data/number-format-cases.json'), 'utf8'));

export const MAIN_NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
export const REL_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const RELS_NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const relsXml = (items: Array<[id: string, kind: string, target: string, mode?: string]>): string =>
  `<Relationships ${RELS_NS}>${items
    .map(([id, kind, target, mode]) => `<Relationship Id="${id}" Type="${kind.includes('://') ? kind : `${OD}/${kind}`}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/>`)
    .join('')}</Relationships>`;

/** A zip of `files` (text or bytes), with `[Content_Types].xml` added when absent, as the bytes `XlsxPackage.open` takes. */
export async function zipOf(files: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  const zip = new JSZip();
  if (!('[Content_Types].xml' in files))
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  for (const [name, content] of Object.entries(files)) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array' });
}

/** A minimal workbook package: the given workbook body, workbook relationships and extra parts. */
export const workbookZip = (workbook: string, rels: Array<[string, string, string, string?]>, extra: Record<string, string | Uint8Array> = {}) =>
  zipOf({
    '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']]),
    'xl/workbook.xml': `<workbook ${MAIN_NS} ${REL_NS}>${workbook}</workbook>`,
    'xl/_rels/workbook.xml.rels': relsXml(rels),
    ...extra,
  });

/** An element tree from XML text (throws when it is not XML, so a typo in a test fails loudly). */
export function el(xml: string): XmlElement {
  const root = parseXml(xml);
  if (!root) throw new Error(`not XML: ${xml}`);
  return root;
}
