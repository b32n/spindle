import JSZip from 'jszip';
import { MAIN_NS, REL_NS, relsXml } from '../test-support';

// ---------------------------------------------------------------------------
// The hostile packages found by review, as the smallest zips that make them: each is a few hundred KB to a few MB of file that asks
// a reader for gigabytes. `buildShape` makes the zip (in the test process, which has the heap for it); `probe.ts` imports it in a child
// process with a small heap limit, so "bounded memory" is a fact the process proves by surviving, not an estimate.
// ---------------------------------------------------------------------------

const rep = (n: number, f: (i: number) => string): string => {
  const parts: string[] = new Array(n);
  for (let i = 0; i < n; i++) parts[i] = f(i);
  return parts.join('');
};

const CT = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>';
const sheet = (body: string) => `<worksheet ${MAIN_NS}><sheetData>${body}</sheetData></worksheet>`;
const workbook = (sheets: string, extra = '') => `<workbook ${MAIN_NS} ${REL_NS}><sheets>${sheets}</sheets>${extra}</workbook>`;

/** The parts of a one-sheet package, which a shape replaces. */
function basePackage(): Parts {
  return {
    '[Content_Types].xml': `${CT}</Types>`,
    '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']]),
    'xl/workbook.xml': workbook('<sheet name="S" sheetId="1" r:id="rId1"/>'),
    'xl/_rels/workbook.xml.rels': relsXml([['rId1', 'worksheet', 'worksheets/sheet1.xml']]),
    'xl/worksheets/sheet1.xml': sheet('<row r="1"><c r="A1"><v>1</v></c></row>'),
  };
}

const styles = (body: string) => ({ 'xl/styles.xml': `<styleSheet ${MAIN_NS}>${body}</styleSheet>` });
const strings = (body: string) => ({ 'xl/sharedStrings.xml': `<sst ${MAIN_NS}>${body}</sst>` });

/**
 * Entity-dense text of about `chars` characters: units of `&amp;&#NN;&lt;x` (14 characters, three references), the number varying
 * through 26 letters and a counter so that the part compresses by less than the ratio limit of 200:1 (a more regular one is refused
 * before it is read, which is the other defence).
 */
const entityText = (chars: number): string => rep(Math.ceil(chars / 14), (i) => `&amp;&#${65 + (i % 26)};&lt;${(i >> 5) % 10}`);

/** The parts of a package by name: text, or bytes where one part is shared by many names (it is made, and held, once). */
type Parts = Record<string, string | Uint8Array>;

const bytesOf = (text: string): Uint8Array => new TextEncoder().encode(text);

/** A package of `n` sheets, each the text `part(i)` (1-based), with the workbook and its relationships to match; `extra` adds the parts of each sheet (its relationships). */
function manySheets(n: number, part: (i: number) => string | Uint8Array, extra: (i: number) => Record<string, string> = () => ({})): Parts {
  const parts: Parts = {};
  const rels: Array<[string, string, string]> = [];
  for (let i = 1; i <= n; i++) {
    parts[`xl/worksheets/sheet${i}.xml`] = part(i);
    Object.assign(parts, extra(i));
    rels.push([`rId${i}`, 'worksheet', `worksheets/sheet${i}.xml`]);
  }
  return {
    ...parts,
    'xl/workbook.xml': workbook(rels.map(([id], i) => `<sheet name="S${i + 1}" sheetId="${i + 1}" r:id="${id}"/>`).join('')),
    'xl/_rels/workbook.xml.rels': relsXml(rels),
  };
}

/** A relationships part of `n` entries of `kind`, each to a part of its own. */
const relsOfKind = (n: number, kind: string): string => relsXml(Array.from({ length: n }, (_, i) => [`x${i}`, kind, `t${i}.xml`] as [string, string, string]));

/** A number format code of `chars` characters (`0\\$0\\$` repeated: every escape is a plain literal, so the whole code is scanned and copied when it is read). */
const longCode = (chars: number): string => '0\\$0\\$'.repeat(Math.ceil(chars / 6)).slice(0, chars);

/** A styles part with one number format of `chars` characters (id 164) and `xfs` cell formats that all use it. */
const longCodeStyles = (chars: number, xfs: number) =>
  styles(`<numFmts count="1"><numFmt numFmtId="164" formatCode="${longCode(chars).replace(/"/g, '&quot;')}"/></numFmts><cellXfs count="${xfs}">${'<xf numFmtId="164"/>'.repeat(xfs)}</cellXfs>`);

/** The cells A1.. of a column, each with its own style index (`xfs` of them). */
const styledCells = (n: number): string => rep(n, (i) => `<row r="${i + 1}"><c r="A${i + 1}" s="${i}"><v>${i}</v></c></row>`);

/**
 * `chars` characters of filler that compress about 25 to 1, as real text does and unlike a single repeated character (which the compression
 * ratio limit refuses): runs of `x` with a pseudo-random letter every 16th character (a fixed generator, so the file is the same every time).
 */
function filler(chars: number): string {
  let seed = 7;
  const blocks: string[] = new Array(Math.ceil(chars / 16));
  for (let i = 0; i < blocks.length; i++) {
    seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff;
    blocks[i] = `xxxxxxxxxxxxxxx${String.fromCharCode(97 + (seed % 26))}`;
  }
  return blocks.join('').slice(0, chars);
}

/** A part of filler as a comment (the scanner skips it), with one character past Latin-1 so that the decoded text takes two bytes a character. */
const comment = (chars: number): string => `<!--${filler(chars)}\u20ac-->`;

/**
 * `rows` rows of `perRow` empty cells, each `<c/>` or `<c />` by a fixed pseudo-random choice (so that the part compresses about 30 to 1,
 * as a real one does, and is not refused by the ratio limit). The rows and cells have no address: each follows the one before.
 */
function emptyCells(rows: number, perRow: number): string {
  let seed = 11;
  const items: string[] = new Array(rows);
  for (let r = 0; r < rows; r++) {
    let row = '<row>';
    for (let c = 0; c < perRow; c++) {
      seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff;
      row += (seed >> 16) & 1 ? '<c/>' : '<c />';
    }
    items[r] = `${row}</row>`;
  }
  return items.join('');
}

/** `rows` rows of 15 numbers each (columns A to O), so a sheet of `rows` x 15 cells has the shape of a real one, within the sheet's 1,048,576 rows. */
const numberRows = (rows: number): string => rep(rows, (r) => `<row r="${r + 1}">${rep(15, (c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}"><v>${r * 15 + c}</v></c>`)}</row>`);

export const SHAPES = {
  // shared strings
  'si-5m-tiny': () => strings(rep(5_000_000, (i) => `<si><t>${i % 99_991}</t></si>`)),
  'si-2m': () => strings(rep(2_000_000, (i) => `<si><t>string ${i}</t></si>`)),
  'runs-400k': () => strings(`<si>${'<r><rPr><b/></rPr><t>x</t></r>'.repeat(400_000)}</si>`),
  // styles
  'border-1m': () => styles(`<borders>${rep(1_000_000, (i) => `<border><left style="thin"><color rgb="FF${(i % 0xffffff).toString(16).padStart(6, '0')}"/></left><right/><top/><bottom/><diagonal/></border>`)}</borders>`),
  'fill-1m': () => styles(`<fills>${rep(1_000_000, (i) => `<fill><patternFill patternType="solid"><fgColor rgb="FF${(i % 0xffffff).toString(16).padStart(6, '0')}"/></patternFill></fill>`)}</fills>`),
  'font-1m': () => styles(`<fonts>${rep(1_000_000, (i) => `<font><sz val="${i % 100}"/><name val="F${i}"/></font>`)}</fonts>`),
  'dxf-500k': () => styles(`<dxfs>${rep(500_000, (i) => `<dxf><font><b/><color rgb="FF${(i % 0xffffff).toString(16).padStart(6, '0')}"/></font><fill><patternFill><bgColor rgb="FF0000FF"/></patternFill></fill></dxf>`)}</dxfs>`),
  // theme, workbook, relationships, content types
  'theme-1m': () => ({ 'xl/theme/theme1.xml': `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme>${rep(1_000_000, (i) => `<a:accent${(i % 6) + 1}><a:srgbClr val="${(i % 0xffffff).toString(16).padStart(6, '0')}"/></a:accent${(i % 6) + 1}>`)}</a:clrScheme></a:themeElements></a:theme>` }),
  'sheet-rels-1m': () => ({ 'xl/_rels/workbook.xml.rels': relsXml([['rId1', 'worksheet', 'worksheets/sheet1.xml'], ...Array.from({ length: 1_000_000 }, (_, i) => [`x${i}`, 'worksheet', `worksheets/s${i}.xml`] as [string, string, string])]) }),
  'defined-names-1m': () => ({ 'xl/workbook.xml': workbook('<sheet name="S" sheetId="1" r:id="rId1"/>', `<definedNames>${rep(1_000_000, (i) => `<definedName name="n${i}">S!$A$${(i % 1000) + 1}</definedName>`)}</definedNames>`) }),
  'content-types-1m': () => ({ '[Content_Types].xml': `${CT}${rep(1_000_000, (i) => `<Override PartName="/xl/p${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`)}</Types>` }),
  // entity-dense text, as one string (shared strings, which keep it) and as a cell value (which the reader decodes and drops)
  'entities-100mb': () => strings(`<si><t>${entityText(100_000_000)}</t></si>`),
  'entities-42mb': () => strings(`<si><t>${entityText(42_000_000)}</t></si>`),
  'entities-cell-100mb': () => ({ 'xl/worksheets/sheet1.xml': sheet(`<row r="1"><c r="A1" t="str"><v>${entityText(100_000_000)}</v></c></row>`) }),
  // rows: eight sheets of every row with a height and no cell
  'rows-8x1m': () => {
    const rows = rep(1_048_576, (i) => `<row r="${i + 1}" ht="${15 + (i % 3)}" customHeight="1"/>`);
    const parts: Record<string, string> = {};
    const rels: Array<[string, string, string]> = [];
    for (let s = 1; s <= 8; s++) {
      parts[`xl/worksheets/sheet${s}.xml`] = sheet(rows);
      rels.push([`rId${s}`, 'worksheet', `worksheets/sheet${s}.xml`]);
    }
    return {
      ...parts,
      'xl/workbook.xml': workbook(rels.map(([id], i) => `<sheet name="S${i + 1}" sheetId="${i + 1}" r:id="${id}"/>`).join('')),
      'xl/_rels/workbook.xml.rels': relsXml(rels),
    };
  },
  // names and values copied into warnings
  'name-10mb': () => ({ 'xl/workbook.xml': workbook(`<sheet name="${'N'.repeat(10_000_000)}" sheetId="1" r:id="rId9"/>`) }),
  'value-1m-digits': () => ({ 'xl/worksheets/sheet1.xml': sheet(`<row r="1"><c r="A1"><v>${'9'.repeat(1_000_000)}</v></c><c r="B1" t="d"><v>${'9'.repeat(1_000_000)}</v></c></row>`) }),
  // the formats of review round 5: a long code under every format (it was rescanned and copied for each distinct style index)
  'numfmt-1mib-2000xf': () => ({ ...longCodeStyles(1024 * 1024, 2000), 'xl/worksheets/sheet1.xml': sheet(styledCells(2000)) }),
  'numfmt-4mib-900xf': () => ({ ...longCodeStyles(4 * 1024 * 1024, 900), 'xl/worksheets/sheet1.xml': sheet(styledCells(900)) }),
  // strings the result keeps must not keep the part alive: a 13-character dimension of a 120 MB part, and 500 sheets of 1 MB (two bytes a character)
  'dimension-120mb': () => ({ 'xl/worksheets/sheet1.xml': `<worksheet ${MAIN_NS}><dimension ref="A1:XFD1048576"/><sheetData>${comment(120_000_000)}</sheetData></worksheet>` }),
  'sheets-500x1mb': () => {
    const part = bytesOf(`<worksheet ${MAIN_NS}><dimension ref="A1:XFD1048576"/><sheetData>${comment(1_000_000)}<row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`);
    return manySheets(500, () => part);
  },
  // package-wide budgets: columns, extension lists, relationships, scanned elements, defined names
  'cols-300x16k': () => {
    const part = bytesOf(`<worksheet ${MAIN_NS}><cols>${rep(16_384, (c) => `<col min="${c + 1}" max="${c + 1}" width="10"/>`)}</cols><sheetData><row r="1"><c r="XFD1"><v>1</v></c></row></sheetData></worksheet>`);
    return manySheets(300, () => part);
  },
  'extlst-12x8mib': () => {
    const part = bytesOf(`<worksheet ${MAIN_NS}><sheetData/><extLst><ext uri="x">${filler(8 * 1024 * 1024 - 100)}</ext></extLst></worksheet>`);
    return manySheets(12, () => part);
  },
  'rels-4x400k-other': () => {
    const rels = relsOfKind(400_000, 'image');
    return manySheets(4, () => sheet(''), (i) => ({ [`xl/worksheets/_rels/sheet${i}.xml.rels`]: rels }));
  },
  'rels-3x400k-links': () => {
    const rels = relsOfKind(400_000, 'hyperlink');
    return manySheets(3, () => sheet(''), (i) => ({ [`xl/worksheets/_rels/sheet${i}.xml.rels`]: rels }));
  },
  'cells-4x1500k': () => {
    const part = bytesOf(sheet(numberRows(100_000)));
    return manySheets(4, () => part);
  },
  'empty-cells-4x5m': () => {
    const part = bytesOf(sheet(emptyCells(500_000, 10)));
    return manySheets(4, () => part);
  },
  'defined-names-150k': () => ({ 'xl/workbook.xml': workbook('<sheet name="S" sheetId="1" r:id="rId1"/>', `<definedNames>${rep(150_000, (i) => `<definedName name="n${i}">S!$A$${(i % 1000) + 1}</definedName>`)}</definedNames>`) }),
} satisfies Record<string, () => Parts>;

export type ShapeName = keyof typeof SHAPES;

const LEVEL_1: ReadonlySet<ShapeName> = new Set<ShapeName>(['dimension-120mb', 'sheets-500x1mb', 'cols-300x16k', 'extlst-12x8mib', 'rels-4x400k-other', 'rels-3x400k-links', 'cells-4x1500k', 'empty-cells-4x5m']);

/** The package of a shape (compressed as a hostile file would be): the base package with the shape's parts in place of its own. */
export async function buildShape(name: ShapeName): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [file, text] of Object.entries({ ...basePackage(), ...SHAPES[name]() })) zip.file(file, text);
  // the big repetitive ones are deflated at level 1: the test spends its time on the reader, not on making the file
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: LEVEL_1.has(name) ? 1 : 6 } });
}
