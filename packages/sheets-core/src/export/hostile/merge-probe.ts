import JSZip from 'jszip';
import { WorkbookImpl } from '../../workbook';
import { importFromXlsx } from '../xlsx';

// Run by `merge-hostile.test.ts` in a child process: a workbook of one sheet with `merges` merged ranges of a shape (argv[2], argv[3]),
// imported by `importFromXlsx`, and what it cost and gave as one JSON line. The test's time limit on the child is the assertion about
// time; nothing is read off this process's clock for it (the figure here is for the PR's table).
const [shape, count] = [process.argv[2], Number(process.argv[3])];
const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const RELS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The `ref` of merge `i`. */
const ref = (i: number): string => {
  // `blocks`: 2 x 2 blocks, one under the other (what a sheet of many merged cells looks like)
  if (shape === 'blocks') return `A${i * 2 + 1}:B${i * 2 + 2}`;
  // `blocksplus`: the blocks, and a last merge over the first two of them (one overlap among many that do not)
  if (shape === 'blocksplus') return i === count - 1 ? 'A1:C4' : `A${i * 2 + 1}:B${i * 2 + 2}`;
  // `overlap`: every merge reaches the next one's rows, with a size that varies (so the part compresses as a real one does)
  if (shape === 'overlap') return `A${i + 1}:${String.fromCharCode(66 + (i % 20))}${i + 2 + (i % 7)}`;
  // `wide`: one merge of the whole sheet over the cells below, and `count` single cells
  return i === 0 ? 'A1:XFD1048576' : `A${i + 1}`;
};

async function main(): Promise<void> {
  // the cells (`count` rows of two cells for `blocks` and `overlap`, so that the merges cover cells), then the merges
  const cells = Array.from({ length: shape === 'wide' ? count : Math.min(count * 2, 100_000) }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c><c r="B${i + 1}"><v>${i}</v></c></row>`).join('');
  const merges = Array.from({ length: count }, (_, i) => `<mergeCell ref="${ref(i)}"/>`).join('');
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships ${RELS}><Relationship Id="rId1" Type="${OD}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook ${MAIN}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships ${RELS}><Relationship Id="rId1" Type="${OD}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet ${MAIN}><sheetData>${cells}</sheetData><mergeCells>${merges}</mergeCells></worksheet>`);
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 1 } });

  const workbook = new WorkbookImpl('wb', 'WB');
  const started = Date.now();
  const { warnings, loss } = await importFromXlsx(bytes, workbook);
  const ms = Date.now() - started;
  const sheet = workbook.getSheet();
  console.log(JSON.stringify({ ms, zipBytes: bytes.length, merges: sheet.getMergedRegions().length, cells: sheet.cells.size, canUndo: workbook.canUndo(), warnings: warnings.slice(0, 5), loss: loss.items.map((i) => [i.id, i.count]) }));
}

main().catch((e: unknown) => {
  console.log(JSON.stringify({ error: String(e) }));
  process.exitCode = 1;
});
