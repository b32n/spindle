import JSZip from 'jszip';
import { WorkbookImpl } from '../../workbook';
import { exportToXlsx } from '../xlsx';

// Run by `formula-hostile.test.ts` in a child process: `count` array formulas typed into one workbook, each spilling `cells` cells
// (argv[2], argv[3]: just under the per-result cap), exported with the options of argv[4] (`maxArrayCells`) and argv[5]
// (`maxTotalArrayCells`), then read back: the cells the file's array formulas cover (their `ref`s: an import would spill them in full again) and the export's warnings, as one JSON line.
const [count, cells, maxArrayCells, maxTotalArrayCells] = process.argv.slice(2, 6).map(Number);

async function main(): Promise<void> {
  const workbook = new WorkbookImpl('wb', 'WB');
  for (let i = 0; i < count; i++) workbook.setFormula(undefined, 0, i, `=SEQUENCE(${cells})`);
  const warnings: string[] = [];
  const exported = await exportToXlsx(workbook, { maxArrayCells, maxTotalArrayCells, onWarning: (m) => warnings.push(m) });
  const xml = await (await JSZip.loadAsync(exported)).file('xl/worksheets/sheet1.xml')!.async('string');
  const col = (letters: string): number => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  let arrayCells = 0;
  for (const m of xml.matchAll(/<f t="array" ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"/g)) {
    arrayCells += (col(m[3] ?? m[1]) - col(m[1]) + 1) * (Number(m[4] ?? m[2]) - Number(m[2]) + 1);
  }
  console.log(JSON.stringify({ exportBytes: exported.length, arrayCells, warnings }));
}

main().catch((e: unknown) => {
  console.log(JSON.stringify({ error: String(e) }));
  process.exitCode = 1;
});
