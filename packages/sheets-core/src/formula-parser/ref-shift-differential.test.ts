// Differential test for S4b: a cached-only cell (text shifted by shiftFormulaRefs) and the equivalent normal formula cell
// (stable-id AST, re-rendered) must hold the same formula TEXT after the same structural edits.

import { WorkbookImpl } from '../workbook';
import { renderStableAst } from './stable-ast';
import type { StructuralEdit } from './ref-shift';

// A small deterministic PRNG (mulberry32): the cases are the same on every run.
function prng(seed: number): () => number {
  let a = seed;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const SHEETS = ['Main', 'Other', 'My Sheet'];
const PREFIXES = ['', '', '', 'Other!', "'My Sheet'!", 'Main!', 'Ghost!'];
const COLS = 'ABCDEFGH';

const pick = <T,>(rand: () => number, items: readonly T[]): T => items[Math.floor(rand() * items.length)];
const int = (rand: () => number, max: number): number => Math.floor(rand() * max);

function cell(rand: () => number): string {
  return `${rand() < 0.3 ? '$' : ''}${pick(rand, COLS.split(''))}${rand() < 0.3 ? '$' : ''}${1 + int(rand, 10)}`;
}

function operand(rand: () => number): string {
  const prefix = pick(rand, PREFIXES);
  const kind = int(rand, 8);
  if (kind <= 2) return `${prefix}${cell(rand)}`;
  if (kind <= 4) return `${prefix}${cell(rand)}:${cell(rand)}`;
  if (kind === 5) return `${prefix}${rand() < 0.3 ? '$' : ''}${pick(rand, COLS.split(''))}:${rand() < 0.3 ? '$' : ''}${pick(rand, COLS.split(''))}`;
  if (kind === 6) return `${prefix}${rand() < 0.3 ? '$' : ''}${1 + int(rand, 10)}:${rand() < 0.3 ? '$' : ''}${1 + int(rand, 10)}`;
  return pick(rand, ["Other:'My Sheet'!A5", `${cell(rand)}#`, '"A1 in a string"', '42', 'TRUE']);
}

function formula(rand: () => number): string {
  const parts = Array.from({ length: 1 + int(rand, 3) }, () => operand(rand));
  const fn = pick(rand, ['SUM', 'NOSUCHFN', 'MAX', 'COUNT']);
  const body = parts.length === 1 && rand() < 0.4 ? parts[0] : `${fn}(${parts.join(',')})`;
  return rand() < 0.3 ? `=${body}+${operand(rand)}` : `=${body}`;
}

function edit(rand: () => number): StructuralEdit {
  return { axis: pick(rand, ['row', 'col'] as const), kind: pick(rand, ['insert', 'delete'] as const), start: int(rand, 12), count: 1 + int(rand, 3) };
}

function applyEdit(wb: WorkbookImpl, sheetName: string, e: StructuralEdit): void {
  const sheet = wb.getSheet(wb.getSheetIdByName(sheetName));
  if (e.axis === 'row' && e.kind === 'insert') sheet.insertRows(e.start, e.count);
  else if (e.axis === 'row') sheet.deleteRows(e.start, e.count);
  else if (e.kind === 'insert') sheet.insertCols(e.start, e.count);
  else sheet.deleteCols(e.start, e.count);
}

/** A workbook with one formula at (20, 20) of `ownerName`, entered as a normal formula or as a cached-only one. */
function build(ownerName: string, text: string, cachedOnly: boolean): WorkbookImpl {
  const wb = new WorkbookImpl('wb', 'Main');
  wb.addSheet('Other');
  wb.addSheet('My Sheet');
  const owner = wb.getSheetIdByName(ownerName);
  if (cachedOnly) wb.setCell(owner, 20, 20, { formula: text, formulaAst: undefined, value: 7, cachedOnly: true });
  else wb.setFormula(owner, 20, 20, text);
  return wb;
}

/**
 * The formula text of every formula cell. A normal formula's text is rendered from its stable AST right now (the cell's own text is
 * only refreshed when its evaluation gets that far, which a formula that evaluates to an error does not always).
 */
function formulaTexts(wb: WorkbookImpl): string[] {
  const resolver = { getSheet: (name?: string) => (name && wb.getSheetIdByName(name) ? wb.getSheet(wb.getSheetIdByName(name)) : undefined) };
  const out: string[] = [];
  for (const id of SHEETS.map((n) => wb.getSheetIdByName(n))) {
    const sheet = wb.getSheet(id);
    for (const [, , c] of sheet.entries()) {
      if (c.formulaAst) out.push(renderStableAst(c.formulaAst, resolver, sheet));
      else if (c.formula) out.push(c.formula);
    }
  }
  return out;
}

describe('cached-only reference shifting agrees with the normal formula path', () => {
  const rand = prng(20260401);
  const cases = Array.from({ length: 6000 }, () => ({
    text: formula(rand),
    owner: pick(rand, SHEETS),
    sheet: pick(rand, SHEETS),
    edits: Array.from({ length: 1 + int(rand, 3) }, () => edit(rand)),
  }));

  it('gives the same text for 6000 random formulas and sequences of 1 to 3 insert/delete row/column edits', () => {
    let compared = 0;
    let moved = 0;
    let deleted = 0;
    for (const c of cases) {
      const normal = build(c.owner, c.text, false);
      const kept = build(c.owner, c.text, true);
      // only formulas the engine writes back as typed can be compared: the generator emits that form
      if (formulaTexts(normal)[0] !== c.text) continue;
      for (const e of c.edits) {
        applyEdit(normal, c.sheet, e);
        applyEdit(kept, c.sheet, e);
        const [a, b] = [formulaTexts(normal), formulaTexts(kept)];
        if (a.join('\n') !== b.join('\n')) throw new Error(`${c.text} on ${c.owner}, edit ${JSON.stringify(e)} of ${c.sheet}: normal ${a[0]} vs cached-only ${b[0]}`);
      }
      compared++;
      const finalText = formulaTexts(kept)[0];
      if (finalText !== c.text) moved++;
      if (finalText?.includes('#REF!')) deleted++;
    }
    // the corpus really exercises the paths: most formulas compared, many moved, many hit deleted rows/columns
    expect(compared).toBeGreaterThan(5000);
    expect(moved).toBeGreaterThan(1000);
    expect(deleted).toBeGreaterThan(200);
  });
});
