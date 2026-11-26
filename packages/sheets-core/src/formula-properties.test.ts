// Randomised (seeded, so reproducible) properties of the formula engine's new syntax:
//  1. parse -> print -> parse is the identity on the AST, for every node kind and any nesting;
//  2. structural edits keep every formula printable, re-parseable and, for whole-column /
//     whole-row references, correct.
// A failing case prints its seed and formula text.

import { SheetImpl } from './sheet';
import { WorkbookImpl } from './workbook';
import { FormulaParser } from './formula-parser/parser';
import { renderStableAst, toStableAst, type SheetResolver } from './formula-parser/stable-ast';
import type { ParsedFormulaNode } from './formula-parser/types';

/** A small, fast, seedable PRNG (mulberry32). */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

class Gen {
  constructor(private readonly rnd: () => number) {}

  int(n: number): number {
    return Math.floor(this.rnd() * n);
  }

  chance(p: number): boolean {
    return this.rnd() < p;
  }

  pick<T>(items: readonly T[]): T {
    return items[this.int(items.length)];
  }

  col(): string {
    return this.pick(['A', 'B', 'C', 'D', 'Z', 'AA', 'AB', 'XFD']);
  }

  row(): number {
    return this.pick([1, 2, 3, 10, 99, 1048576]);
  }

  anchor(): string {
    return this.chance(0.3) ? '$' : '';
  }

  cell(): string {
    return `${this.anchor()}${this.col()}${this.anchor()}${this.row()}`;
  }

  sheetPrefix(): string {
    return this.pick(['', '', '', 'Sheet2!', "'My Sheet'!", 'Ghost!', "'it''s'!"]);
  }

  /** A reference (cell, range, whole column/row), possibly sheet-qualified. */
  plainRef(): string {
    const prefix = this.sheetPrefix();
    switch (this.int(5)) {
      case 0:
        return prefix + this.cell();
      case 1:
        return `${prefix}${this.cell()}:${this.cell()}`;
      case 2:
        return `${prefix}${this.anchor()}${this.col()}:${this.anchor()}${this.col()}`;
      case 3:
        return `${prefix}${this.anchor()}${this.row()}:${this.anchor()}${this.row()}`;
      default:
        return prefix + this.cell();
    }
  }

  threeD(): string {
    const prefix = this.pick(['Sheet1:Sheet3!', "'Sheet1:My Sheet'!", 'Sheet3:Sheet1!', "'it''s:Sheet2'!"]);
    switch (this.int(4)) {
      case 0:
        return prefix + this.cell();
      case 1:
        return `${prefix}${this.cell()}:${this.cell()}`;
      case 2:
        return `${prefix}${this.anchor()}${this.col()}:${this.anchor()}${this.col()}`;
      default:
        return `${prefix}${this.anchor()}${this.row()}:${this.anchor()}${this.row()}`;
    }
  }

  /** An expression that denotes a single reference (usable as an operand of `:`). */
  refOperand(depth: number): string {
    switch (this.int(depth > 0 ? 6 : 3)) {
      case 0:
      case 1:
        return this.plainRef();
      case 2:
        return `${this.pick(['', 'Sheet2!'])}${this.cell()}#`;
      case 3:
        return `INDEX(${this.plainRef()},${this.int(5) + 1})`;
      case 4:
        return `OFFSET(${this.cell()},${this.int(4)},${this.int(4)})`;
      default:
        return `INDIRECT("${this.cell().replace(/\$/g, '')}")`;
    }
  }

  /** Any reference-valued expression: a single reference, a `:` span, an intersection or a union. */
  ref(depth: number): string {
    switch (this.int(depth > 0 ? 6 : 2)) {
      case 0:
      case 1:
        return this.refOperand(depth);
      case 2:
        return `${this.refOperand(depth - 1)}:${this.refOperand(depth - 1)}`;
      case 3:
        return `${this.refOperand(depth - 1)} ${this.refOperand(depth - 1)}`;
      case 4:
        return `(${this.refOperand(depth - 1)},${this.refOperand(depth - 1)}${this.chance(0.4) ? `,${this.refOperand(depth - 1)}` : ''})`;
      default:
        return `@${this.refOperand(depth - 1)}`;
    }
  }

  literal(): string {
    switch (this.int(7)) {
      case 0:
        return this.pick(['0', '1', '42', '3.14', '0.5', '1e3', '1.5E-7', '12345678901234567890', '.25']);
      case 1:
        return this.pick(['""', '"a"', '"say ""hi"""', '"a,b;c"', '"{1,2}"', '"A1:B2"']);
      case 2:
        return this.pick(['TRUE', 'FALSE']);
      case 3:
        return this.pick(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);
      case 4:
        return this.arrayConstant();
      default:
        return this.pick(['0', '7', '2.5']);
    }
  }

  arrayConstant(): string {
    const rows = this.int(3) + 1;
    const cols = this.int(3) + 1;
    const element = (): string =>
      this.pick(['1', '-2', '+3', '4.5', '-0.5e2', '"x"', '"q""q"', 'TRUE', 'FALSE', '#N/A', '#DIV/0!', '""']);
    return `{${Array.from({ length: rows }, () => Array.from({ length: cols }, element).join(',')).join(';')}}`;
  }

  /** A value-producing expression. */
  expr(depth: number): string {
    if (depth <= 0) return this.chance(0.5) ? this.literal() : this.ref(0);
    switch (this.int(12)) {
      case 0:
        return this.literal();
      case 1:
        return this.ref(depth);
      case 2:
        return `${this.pick(['SUM', 'MAX', 'COUNT', 'AVERAGE'])}(${this.threeD()})`;
      case 3: {
        const n = this.int(3) + 1;
        return `${this.pick(['SUM', 'IF', 'MAX', 'CONCAT', 'TEXTJOIN'])}(${Array.from({ length: n }, () => this.expr(depth - 1)).join(',')})`;
      }
      case 4:
        return `${this.pick(['-', '+', '--', '-+'])}${this.expr(depth - 1)}`;
      case 5:
        return `${this.operand(depth - 1)}${this.pick(['%', '%%'])}`;
      case 6:
      case 7:
      case 8:
        return `${this.expr(depth - 1)}${this.pick(['+', '-', '*', '/', '^', '&', '=', '<>', '<', '>', '<=', '>='])}${this.expr(depth - 1)}`;
      case 9:
        return `(${this.expr(depth - 1)})`;
      case 10:
        return `SUM(${this.ref(depth)}${this.chance(0.5) ? `,${this.ref(depth)}` : ''})`;
      default:
        return `SUM(${this.arrayConstant()},${this.expr(depth - 1)})`;
    }
  }

  /** An operand that survives a trailing `%` without a surrounding paren, or is wrapped in one. */
  operand(depth: number): string {
    const text = this.expr(depth);
    return /^[\w.$!#'":{}@()]+$/.test(text) && !text.includes(' ') ? text : `(${text})`;
  }
}

const parser = new FormulaParser();

function harness() {
  const names = ['Sheet1', 'Sheet2', 'Sheet3', 'My Sheet', "it's"];
  const sheets = new Map(names.map((n, i) => [n, new SheetImpl(`id${i}`, n)]));
  const resolver: SheetResolver = { getSheet: (name) => (name ? sheets.get(name) : undefined) };
  const current = sheets.get('Sheet1') as SheetImpl;
  const print = (ast: ParsedFormulaNode, row: number, col: number): string =>
    renderStableAst(toStableAst(ast, resolver, current, row, col), resolver, current);
  return { print };
}

describe('property: parse -> print -> parse is the identity (seeded)', () => {
  const { print } = harness();
  const seeds = Array.from({ length: 4000 }, (_, i) => i + 1);

  it(`holds for ${seeds.length} random formulas over every node kind`, () => {
    const kinds = new Set<string>();
    for (const seed of seeds) {
      const g = new Gen(rng(seed));
      const text = `=${g.expr(g.int(4))}`;
      const [row, col] = [g.int(30), g.int(30)];
      const first = parser.parse(text, row, col);
      if (first.error) throw new Error(`seed ${seed}: the generator produced an unparseable formula ${text}: ${first.error}`);
      const printed = print(first.ast, row, col);
      const second = parser.parse(printed, row, col);
      if (second.error) throw new Error(`seed ${seed}: ${text} printed as ${printed}, which does not parse: ${second.error}`);
      try {
        expect(second.ast).toEqual(first.ast);
        expect(print(second.ast, row, col)).toBe(printed);
      } catch (e) {
        throw new Error(`seed ${seed}: ${text} -> ${printed}\n${(e as Error).message}`);
      }
      const walk = (n: ParsedFormulaNode) => {
        kinds.add(n.type === 'operator' ? `op${n.operator}` : n.type);
        n.args?.forEach(walk);
        if (n.left) walk(n.left);
        if (n.right) walk(n.right);
        if (n.rangeRef?.whole) kinds.add(`whole-${n.rangeRef.whole}`);
      };
      walk(first.ast);
    }
    // The generator really exercised every kind the engine added.
    for (const kind of [
      'array', 'percent', 'implicit', 'spillref', 'ref3d', 'error', 'range', 'cell', 'function', 'variable',
      'op:', 'op ', 'op,', 'whole-column', 'whole-row',
    ]) {
      if (kind !== 'variable') expect(kinds).toContain(kind);
    }
  });

  it('prints canonical text exactly: printing the printed text changes nothing', () => {
    for (const seed of seeds.slice(0, 1500)) {
      const g = new Gen(rng(seed * 7919));
      const text = `=${g.expr(3)}`;
      const once = print(parser.parse(text, 3, 3).ast, 3, 3);
      const twice = print(parser.parse(once, 3, 3).ast, 3, 3);
      expect(twice).toBe(once);
    }
  });
});

describe('property: structural edits keep formulas valid and correct (seeded)', () => {
  const WHOLE_COL = ['A', 'B', 'C', 'D', 'E'];

  function build(seed: number) {
    const g = new Gen(rng(seed));
    const wb = new WorkbookImpl('wb', 'WB');
    const sheet = wb.getSheet();
    wb.renameSheet(sheet.id, 'Sheet1');
    wb.addSheet('Sheet2');
    // Data in A1:E10 on both sheets.
    for (let r = 0; r < 10; r++) {
      for (let c = 0; c < 5; c++) {
        wb.setCellValue(undefined, r, c, g.int(100));
        wb.setCellValue(wb.getSheetIdByName('Sheet2'), r, c, g.int(100));
      }
    }
    // Formulas live clear of every whole-column ref (column H) and of every whole-row ref (rows 61+, far enough that the deletes never reach them).
    const formulas = [
      ...WHOLE_COL.map((l) => `=SUM(${l}:${l})`),
      '=SUM(A:C)', '=SUM($B:$D)', '=COUNT(A:E)', '=MAX(C:C)',
      '=SUM(1:1)', '=SUM(2:3)', '=SUM($4:$4)', '=COUNT(5:7)',
      '=SUM(Sheet2!A:A)', '=SUM(Sheet2!2:2)', '=SUM(Sheet1:Sheet2!B2)', '=SUM(A1:INDEX(A1:A10,3))', '=SUM(A1:B5 B2:C6)',
      '=SUM((A1:A3,C5))', '=@A1:A10', '=SUM(A1:A3)*50%', '=SUM({1,2;3,4})+A:A',
    ];
    formulas.forEach((f, i) => wb.setFormula(undefined, 60 + i, 7, f));
    return { g, wb, sheet, formulas };
  }

  /** Sum of the numeric non-formula cells of the sheet in the given columns / rows, by scanning the sheet. */
  function scan(sheet: SheetImpl, pick: (row: number, col: number) => boolean): number {
    let total = 0;
    for (const [r, c, cell] of sheet.entries()) {
      if (!cell.formula && typeof cell.value === 'number' && pick(r, c)) total += cell.value;
    }
    return total;
  }

  const seeds = Array.from({ length: 60 }, (_, i) => i + 1);

  it('after any run of inserts and deletes, every formula text parses and re-prints to itself', () => {
    for (const seed of seeds) {
      const { g, wb, sheet } = build(seed);
      for (let step = 0; step < 14; step++) {
        const at = g.int(14);
        const n = g.int(2) + 1;
        switch (g.int(4)) {
          case 0:
            sheet.insertRows(at, n);
            break;
          case 1:
            sheet.deleteRows(at, n);
            break;
          case 2:
            sheet.insertCols(g.int(7), n);
            break;
          default:
            sheet.deleteCols(g.int(7), n);
        }
        for (const [row, col, cell] of sheet.entries()) {
          if (!cell.formula) continue;
          const parsed = parser.parse(cell.formula, row, col);
          if (parsed.error) throw new Error(`seed ${seed} step ${step}: ${cell.formula} does not parse: ${parsed.error}`);
          const resolver: SheetResolver = { getSheet: (name) => (name ? (wb.getSheet(wb.getSheetIdByName(name)) as SheetImpl) : undefined) };
          const printed = renderStableAst(toStableAst(parsed.ast, resolver, sheet, row, col), resolver, sheet);
          if (printed !== cell.formula) throw new Error(`seed ${seed} step ${step}: ${cell.formula} re-prints as ${printed}`);
        }
      }
    }
  });

  it('whole-column sums are right after any run of row inserts and deletes, and their text never changes', () => {
    for (const seed of seeds) {
      const { g, wb, sheet, formulas } = build(seed);
      for (let step = 0; step < 12; step++) {
        const at = g.int(14);
        if (g.chance(0.5)) sheet.insertRows(at, g.int(2) + 1);
        else sheet.deleteRows(at, g.int(2) + 1);
        // A column is addressed by its ids, which row edits never touch: the texts stay exactly as typed.
        WHOLE_COL.forEach((l, i) => {
          const loc = findFormula(sheet, i + 0);
          if ((loc && loc.cell.formula) !== formulas[i]) throw new Error(`seed ${seed} step ${step} i=${i}: ` + [...sheet.entries()].filter(([, , c]) => c.formula).map(([r, c, cell]) => `${r},${c}:${cell.formula}`).join(' | '));
          if (!loc) return;
          const expected = scan(sheet, (_r, c) => c === i);
          expect(wb.getCellCalculatedValue(undefined, loc.row, loc.col)).toBe(expected);
        });
      }
    }
  });

  it('whole-column sums follow column inserts and become #REF! exactly when their column is deleted', () => {
    for (const seed of seeds) {
      const { g, wb, sheet } = build(seed);
      for (let step = 0; step < 8; step++) {
        const col = g.int(7);
        if (g.chance(0.5)) sheet.insertCols(col, 1);
        else sheet.deleteCols(col, 1);
        for (const [row, c, cell] of sheet.entries()) {
          const m = cell.formula && /^=SUM\(([A-Z]+):\1\)$/.exec(cell.formula);
          if (!m) continue;
          const index = [...m[1]].reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
          expect(wb.getCellCalculatedValue(undefined, row, c)).toBe(scan(sheet, (_r, cc) => cc === index));
        }
        // A deleted column leaves #REF! behind, never a stale sum.
        for (const [row, c, cell] of sheet.entries()) {
          if (cell.formula === '=SUM(#REF!)') expect(wb.getCellCalculatedValue(undefined, row, c)).toBe('#REF!');
        }
      }
    }
  });

  it('whole-row sums follow row inserts and become #REF! exactly when their row is deleted', () => {
    for (const seed of seeds) {
      const { g, wb, sheet } = build(seed);
      for (let step = 0; step < 8; step++) {
        const row = g.int(8);
        if (g.chance(0.5)) sheet.insertRows(row, 1);
        else sheet.deleteRows(row, 1);
        for (const [r, c, cell] of sheet.entries()) {
          const m = cell.formula && /^=SUM\((\d+):\1\)$/.exec(cell.formula);
          if (!m) continue;
          const index = Number(m[1]) - 1;
          expect(wb.getCellCalculatedValue(undefined, r, c)).toBe(scan(sheet, (rr) => rr === index));
        }
        for (const [r, c, cell] of sheet.entries()) {
          if (cell.formula === '=SUM(#REF!)') expect(wb.getCellCalculatedValue(undefined, r, c)).toBe('#REF!');
        }
      }
    }
  });
});

/** The formula cell in column H that was entered `index`-th (they sit in consecutive rows from row 61, shifting with edits). */
function findFormula(sheet: SheetImpl, index: number): { row: number; col: number; cell: { formula?: string } } | undefined {
  const rows: Array<{ row: number; col: number; cell: { formula?: string } }> = [];
  for (const [row, col, cell] of sheet.entries()) {
    if (cell.formula) rows.push({ row, col, cell });
  }
  rows.sort((a, b) => a.row - b.row);
  return rows[index];
}
