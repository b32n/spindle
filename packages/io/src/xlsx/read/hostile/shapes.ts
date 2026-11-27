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
function basePackage(): Record<string, string> {
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
} satisfies Record<string, () => Record<string, string>>;

export type ShapeName = keyof typeof SHAPES;

/** The package of a shape (compressed as a hostile file would be): the base package with the shape's parts in place of its own. */
export async function buildShape(name: ShapeName): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [file, text] of Object.entries({ ...basePackage(), ...SHAPES[name]() })) zip.file(file, text);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
