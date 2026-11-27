import JSZip from 'jszip';
import { WorkbookImpl } from '../../workbook';
import { importFromXlsx } from '../xlsx';

// Run by `formula-hostile.test.ts` in a child process: a one-sheet workbook whose formulas have a hostile shape (argv[2], argv[3]),
// imported by `importFromXlsx`, and what it cost and gave as one JSON line. The test's heap and time limits on the child are the
// assertions about memory and time; nothing is read off this process's clock for them (the figure here is for the PR's table).
const [shape, count] = [process.argv[2], Number(process.argv[3])];
const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const RELS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');

/** The `<c>` of formula cell `i`, by shape. */
function cell(i: number): string {
  const ref = `A${i + 1}`;
  switch (shape) {
    // array formulas over a block of 26,000 cells each (a legacy one keeps its extent, a dynamic one clears what its saved result occupies)
    case 'arrays-legacy':
      return `<c r="${ref}"><f t="array" ref="A1:Z1000">SUM(1,2)</f><v>3</v></c>`;
    case 'arrays-dynamic':
      return `<c r="${ref}" cm="1"><f t="array" ref="A1:Z1000">SUM(1,2)</f><v>3</v></c>`;
    // one array formula over the whole sheet
    case 'array-huge':
      return `<c r="${ref}"><f t="array" ref="A1:XFD1048576">SUM(1,2)</f><v>3</v></c>`;
    // 8,000 characters of nesting: parsed or refused, never a stack overflow
    case 'nested':
      return `<c r="${ref}"><f>${esc('('.repeat(4000) + '1' + ')'.repeat(4000))}</f><v>1</v></c>`;
    case 'singles':
      return `<c r="${ref}"><f>${'_xlfn.SINGLE('.repeat(600)}A1${')'.repeat(600)}</f><v>1</v></c>`;
    // 450 distinct function names a formula (a formula over 8,192 characters is dropped by the reader): every one is a name the detector would list
    case 'names':
      return `<c r="${ref}"><f>${Array.from({ length: 450 }, (_, k) => `FN${i}X${k}(1)`).join('+')}</f><v>1</v></c>`;
    default:
      throw new Error(`unknown shape ${shape}`);
  }
}

async function main(): Promise<void> {
  const rows = Array.from({ length: count }, (_, i) => `<row r="${i + 1}">${cell(i)}</row>`).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships ${RELS}><Relationship Id="rId1" Type="${OD}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook ${MAIN}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships ${RELS}><Relationship Id="rId1" Type="${OD}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet ${MAIN}><sheetData>${rows}</sheetData></worksheet>`);
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 1 } });

  const workbook = new WorkbookImpl('wb', 'WB');
  const started = Date.now();
  const { warnings, loss } = await importFromXlsx(bytes, workbook);
  const ms = Date.now() - started;
  const sheet = workbook.getSheet();
  let cachedOnly = 0;
  for (const [, , c] of sheet.entries()) if (c.cachedOnly) cachedOnly++;
  const names = loss.items.find((i) => i.id === 'xlsx.formulas.unsupported');
  console.log(
    JSON.stringify({
      ms,
      zipBytes: bytes.length,
      cells: sheet.cells.size,
      cachedOnly,
      warnings: warnings.slice(0, 5),
      loss: loss.items.map((i) => [i.id, i.count]),
      unsupportedNames: names?.where?.length ?? 0,
    })
  );
}

main().catch((e: unknown) => {
  console.log(JSON.stringify({ error: String(e) }));
  process.exitCode = 1;
});
