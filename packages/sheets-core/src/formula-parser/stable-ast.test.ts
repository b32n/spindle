import { SheetImpl } from '../sheet';
import { FormulaParser } from './parser';
import {
  StableRefDeletedError,
  WHOLE_AXIS_ID,
  adjustStableAstForCopy,
  collectStableDependencies,
  formulaIsVolatile,
  fromStableAst,
  renameSheetInAst,
  renderStableAst,
  shrinkThreeDSpansForDelete,
  toStableAst,
  walkStableAst,
  type SheetResolver,
  type StableFormulaNode,
} from './stable-ast';

const parser = new FormulaParser();

function setup(names = ['Sheet1', 'Sheet2', 'Sheet3']) {
  const sheets = new Map(names.map((n, i) => [n, new SheetImpl(`id${i}`, n)]));
  const resolver: SheetResolver = { getSheet: (name) => (name ? sheets.get(name) : undefined) };
  const current = sheets.get(names[0]) as SheetImpl;
  const stable = (formula: string, row = 0, col = 0): StableFormulaNode =>
    toStableAst(parser.parse(formula, row, col).ast, resolver, current, row, col);
  const text = (node: StableFormulaNode): string => renderStableAst(node, resolver, current);
  const back = (node: StableFormulaNode, row = 0, col = 0) => fromStableAst(node, resolver, current, row, col);
  return { sheets, resolver, current, stable, text, back };
}

describe('numeric <-> stable conversion', () => {
  it.each([
    '=1+2*3', '="a"&"b"', '=TRUE', '=#NULL!', '=Name1', '=A1', '=$B$2+C3', '=A1:B2', '=SUM(A:A)', '=SUM($B:$D)', '=SUM(1:1)',
    '=SUM($2:3)', '=A1%', '=-A1%', '={1,2;3,4}', '=SUM(Sheet2!A1:B2)', '=SUM(Sheet2!A:A)', '=A1:INDEX(A1:A5,3)',
    '=A1:B2 B1:C3', '=SUM((A1,B2))', '=@A1:A3', '=A1#', '=Sheet2!B2#',
  ])('%s converts to a stable AST and back unchanged', (formula) => {
    const { stable, back } = setup();
    const at: [number, number] = [4, 4];
    const numeric = parser.parse(formula, ...at).ast;
    expect(back(stable(formula, ...at), ...at)).toEqual(numeric);
  });

  it('binds each bounded axis to ids and the unbounded axis to the sentinel', () => {
    const { stable, current } = setup();
    const col = stable('=SUM(B:D)').args?.[0].rangeRef;
    expect(col?.whole).toBe('column');
    expect([col?.start.rowId, col?.end.rowId]).toEqual([WHOLE_AXIS_ID, WHOLE_AXIS_ID]);
    expect(col?.start.colId).toBe(current.getColId(1));
    expect(col?.end.colId).toBe(current.getColId(3));
    const row = stable('=SUM(2:3)').args?.[0].rangeRef;
    expect(row?.whole).toBe('row');
    expect([row?.start.colId, row?.end.colId]).toEqual([WHOLE_AXIS_ID, WHOLE_AXIS_ID]);
    expect(row?.start.rowId).toBe(current.getRowId(1));
    expect(row?.end.rowId).toBe(current.getRowId(2));
  });

  it('keeps a plain range free of the whole flag', () => {
    const { stable } = setup();
    expect(stable('=SUM(A1:B2)').args?.[0].rangeRef).not.toHaveProperty('whole');
  });

  it('throws #REF! once the bounded axis id is gone, but not when only the other axis changes', () => {
    const { stable, back, current } = setup();
    const col = stable('=SUM(B:B)');
    const row = stable('=SUM(2:2)');
    current.deleteRows(5, 1);
    current.insertRows(0, 3);
    expect(() => back(col)).not.toThrow();
    current.deleteCols(1, 1);
    expect(() => back(col)).toThrow(StableRefDeletedError);
    expect(() => back(row)).not.toThrow();
    current.deleteRows(4, 1); // row 2 has moved to index 4 after the three inserted rows
    expect(() => back(row)).toThrow(StableRefDeletedError);
  });

  it('throws #REF! for a reference to a sheet that no longer exists', () => {
    const { stable, back, sheets } = setup();
    const node = stable('=SUM(Sheet2!A:A)');
    sheets.delete('Sheet2');
    expect(() => back(node)).toThrow(StableRefDeletedError);
  });

  it('keeps a 3D reference positional, with no ids bound', () => {
    const { stable, current } = setup();
    const cell = stable('=SUM(Sheet1:Sheet3!B2)').args?.[0];
    expect(cell?.type).toBe('ref3d');
    expect(cell?.sheetEnd).toBe('Sheet3');
    expect(cell?.cellRef).toMatchObject({ rowId: '#REF!', colId: '#REF!', unresolvedRow: 1, unresolvedCol: 1, sheetName: 'Sheet1' });
    expect(current.getRowId(1)).toBeUndefined(); // nothing was created on any sheet
  });

  it('keeps a reference to a missing sheet positional until the sheet appears', () => {
    const { stable, text, back, sheets, resolver, current } = setup();
    const node = stable('=SUM(Ghost!A:A)+SUM(Ghost!2:2)');
    expect(text(node)).toBe('=SUM(Ghost!A:A)+SUM(Ghost!2:2)');
    expect(() => back(node)).toThrow(StableRefDeletedError);
    sheets.set('Ghost', new SheetImpl('idg', 'Ghost'));
    const numeric = back(node);
    expect(numeric.left?.args?.[0].rangeRef).toMatchObject({
      whole: 'column',
      start: { row: 0, col: 0 },
      end: { row: 1048575, col: 0 },
    });
    expect(numeric.right?.args?.[0].rangeRef).toMatchObject({ whole: 'row', start: { row: 1, col: 0 } });
    expect(resolver.getSheet('Ghost')).toBeDefined();
    expect(current.name).toBe('Sheet1');
  });

  it('refuses a positional reference whose position fell off the sheet', () => {
    const { stable, back } = setup();
    const node = stable('=SUM(Sheet1:Sheet3!A1)');
    const ref = node.args?.[0].cellRef;
    if (ref) ref.unresolvedRow = -1;
    expect(() => back(node)).toThrow(StableRefDeletedError);
    if (ref) {
      ref.unresolvedRow = 0;
      ref.unresolvedCol = -2;
    }
    expect(() => back(node)).toThrow(StableRefDeletedError);
  });

  it('resolves a relative positional reference against the formula cell', () => {
    const { stable, back } = setup();
    const node = stable('=SUM(Sheet1:Sheet3!A1)', 3, 3); // offsets -3, -3
    // The stored position is absolute (A1); a relative ref is re-encoded as an offset from wherever the formula is.
    expect(back(node, 5, 6).args?.[0].cellRef).toMatchObject({ row: -5, col: -6, rowAbsolute: false });
    const abs = stable('=SUM(Sheet1:Sheet3!$A$1)', 3, 3);
    expect(back(abs, 5, 6).args?.[0].cellRef).toMatchObject({ row: 0, col: 0, rowAbsolute: true });
  });
});

describe('rendering', () => {
  it('prints every node kind in its canonical spelling', () => {
    const { stable, text } = setup(['Sheet1', 'Sheet2', 'Sheet3', 'My Sheet']);
    for (const f of [
      '=1.5', '="say ""hi"""', '=TRUE()', '=#NUM!', '=Total', '=A1', '=$A$1:B$2', '=SUM(A:A)', '=SUM($A:C)', '=SUM(1:2)', '=SUM($1:$2)',
      "=SUM('My Sheet'!A:A)", "=SUM('My Sheet'!1:1)", '=A1%', '=(A1+1)%', '=-(A1%)', '=-A1%', '={1,2;3,4}', '={"a",TRUE,#N/A}',
      '=SUM(Sheet1:Sheet3!A1)', "=SUM('Sheet1:My Sheet'!A1:B2)", '=@A1', '=@(1+2)', '=@(A1%)', '=@A1:A3', '=A1#', "='My Sheet'!A1#",
      '=SUM((A1,B1,C1))', '=SUM((A1:A2,(B1,C1)))', '=A1:B2 B1:C3', '=(A1:B2 B1:C3):D4', '=A1:(B1:C3 B2:D4)', '=SUM((A1:B2 B1:C3,D1))',
      '=(A1+1)*2', '=1-(2-3)', '=2^3^2', '=-2^2', '=2^-2', '=1+2&3', '=1=(2=3)', '=A1:INDEX(A1:A5,3)',
    ]) {
      expect(text(stable(f))).toBe(f);
    }
  });

  it('prints a reference whose rows or columns are gone as #REF!', () => {
    const { stable, text, current } = setup();
    const col = stable('=SUM(B:B)+SUM(B2:C3)+Sheet1:Sheet2!A1');
    current.deleteCols(1, 1);
    expect(text(col)).toBe('=SUM(#REF!)+SUM(#REF!)+Sheet1:Sheet2!A1');
    const row = stable('=SUM(5:5)');
    current.deleteRows(4, 1);
    expect(text(row)).toBe('=SUM(#REF!)');
  });

  it('prints a positional reference moved off the sheet as #REF!', () => {
    const { stable, text } = setup();
    const node = stable('=SUM(Sheet1:Sheet3!A1)+SUM(Sheet1:Sheet3!A:B)+SUM(Sheet1:Sheet3!1:2)');
    const flat: StableFormulaNode[] = [];
    walkStableAst(node, (n) => {
      if (n.type === 'ref3d') flat.push(n);
    });
    expect(flat).toHaveLength(3);
    (flat[0].cellRef as { unresolvedRow: number }).unresolvedRow = -1;
    (flat[1].rangeRef?.start as { unresolvedCol: number }).unresolvedCol = -1;
    (flat[2].rangeRef?.start as { unresolvedRow: number }).unresolvedRow = -1;
    expect(text(node)).toBe('=SUM(#REF!)+SUM(#REF!)+SUM(#REF!)');
  });

  it('prints a node of an unknown type as nothing', () => {
    const { text } = setup();
    expect(text({ type: 'bogus' } as unknown as StableFormulaNode)).toBe('=');
  });

  it('quotes a sheet name only when it needs it, and escapes quotes in it', () => {
    const names = ['Sheet1', 'plain_name.v2', 'A1', 'it\'s', '1st', 'TRUE', 'x-y', 'Übersicht'];
    const { stable, text } = setup(names);
    const f = (sheet: string) => `=${sheet}!B2`;
    expect(text(stable(f('plain_name.v2')))).toBe('=plain_name.v2!B2');
    expect(text(stable(f("'A1'")))).toBe("='A1'!B2");
    expect(text(stable(f("'it''s'")))).toBe("='it''s'!B2");
    expect(text(stable(f("'1st'")))).toBe("='1st'!B2");
    expect(text(stable(f("'TRUE'")))).toBe("='TRUE'!B2");
    expect(text(stable(f("'x-y'")))).toBe("='x-y'!B2");
    expect(text(stable(f('Übersicht')))).toBe('=Übersicht!B2');
  });

  it('normalises spacing and case, but keeps meaning', () => {
    const { stable, text } = setup();
    expect(text(stable('= sum( a:a ) + b1 % '))).toBe('=SUM(A:A)+B1%');
    expect(text(stable('=sum(sheet1:sheet3!a1)'))).toBe('=SUM(sheet1:sheet3!A1)');
  });
});

describe('sheet-name edits on references', () => {
  it('walks every node of an AST', () => {
    const { stable } = setup();
    const seen: string[] = [];
    walkStableAst(stable('=SUM(A1,B1)+C1*-D1'), (n) => seen.push(n.type));
    expect(seen).toEqual(['operator', 'function', 'cell', 'cell', 'operator', 'cell', 'operator', 'string', 'cell']);
    const only: string[] = [];
    walkStableAst({ type: 'number', value: 1 }, (n) => only.push(n.type));
    expect(only).toEqual(['number']);
    walkStableAst({ type: 'operator', operator: '+', left: { type: 'number', value: 1 } }, (n) => only.push(n.type));
    walkStableAst({ type: 'operator', operator: '+', right: { type: 'number', value: 2 } }, (n) => only.push(n.type));
    expect(only).toEqual(['number', 'operator', 'number', 'operator', 'number']);
  });

  it('renames a sheet in cells, ranges and 3D spans', () => {
    const { stable, text, sheets } = setup();
    const node = stable('=Sheet2!A1+SUM(Sheet2!B2:C3)+SUM(Sheet1:Sheet2!A1)+SUM(Sheet2:Sheet3!A1:B2)+Sheet3!A1');
    sheets.set('Renamed', sheets.get('Sheet2') as SheetImpl);
    sheets.delete('Sheet2');
    renameSheetInAst(node, 'Sheet2', 'Renamed');
    expect(text(node)).toBe(
      '=Renamed!A1+SUM(Renamed!B2:C3)+SUM(Sheet1:Renamed!A1)+SUM(Renamed:Sheet3!A1:B2)+Sheet3!A1'
    );
  });

  describe('shrinking a 3D span when a sheet is deleted', () => {
    const order = ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4'];
    const shrink = (formula: string, deleted: string) => {
      const { stable, text } = setup(order);
      const node = stable(formula);
      shrinkThreeDSpansForDelete(node, deleted, order);
      return text(node);
    };

    it('pulls the first end inward', () => {
      expect(shrink('=SUM(Sheet1:Sheet3!A1)', 'Sheet1')).toBe('=SUM(Sheet2:Sheet3!A1)');
      expect(shrink('=SUM(Sheet1:Sheet3!A1:B2)', 'Sheet1')).toBe('=SUM(Sheet2:Sheet3!A1:B2)');
      expect(shrink('=SUM(Sheet1:Sheet3!A:A)', 'Sheet1')).toBe('=SUM(Sheet2:Sheet3!A:A)');
    });

    it('pulls the last end inward', () => {
      expect(shrink('=SUM(Sheet1:Sheet3!A1)', 'Sheet3')).toBe('=SUM(Sheet1:Sheet2!A1)');
      expect(shrink('=SUM(Sheet1:Sheet3!A1:B2)', 'Sheet3')).toBe('=SUM(Sheet1:Sheet2!A1:B2)');
    });

    it('pulls inward for a reversed span', () => {
      expect(shrink('=SUM(Sheet3:Sheet1!A1)', 'Sheet3')).toBe('=SUM(Sheet2:Sheet1!A1)');
      expect(shrink('=SUM(Sheet3:Sheet1!A1)', 'Sheet1')).toBe('=SUM(Sheet3:Sheet2!A1)');
    });

    it('leaves a span alone when an inside or outside sheet is deleted', () => {
      expect(shrink('=SUM(Sheet1:Sheet3!A1)', 'Sheet2')).toBe('=SUM(Sheet1:Sheet3!A1)');
      expect(shrink('=SUM(Sheet1:Sheet3!A1)', 'Sheet4')).toBe('=SUM(Sheet1:Sheet3!A1)');
    });

    it('leaves a one-sheet span and every other node alone', () => {
      expect(shrink('=SUM(Sheet2:Sheet2!A1)', 'Sheet2')).toBe('=SUM(Sheet2:Sheet2!A1)');
      expect(shrink('=SUM(Sheet2!A1)+1', 'Sheet2')).toBe('=SUM(Sheet2!A1)+1');
    });

    it('leaves a span whose sheets are not in the order alone', () => {
      expect(shrink('=SUM(Ghost:Sheet3!A1)', 'Sheet3')).toBe('=SUM(Ghost:Sheet3!A1)');
      expect(shrink('=SUM(Sheet1:Ghost!A1)', 'Sheet1')).toBe('=SUM(Sheet1:Ghost!A1)');
    });
  });
});

describe('copy and fill', () => {
  const rebased = (formula: string, from: [number, number], to: [number, number]) => {
    const { stable, text, resolver, current } = setup();
    const node = stable(formula, ...from);
    return text(adjustStableAstForCopy(node, resolver, current, ...from, ...to));
  };

  it('moves relative parts along and keeps absolute ones', () => {
    expect(rebased('=A1+$B$2', [0, 2], [3, 3])).toBe('=B4+$B$2');
    expect(rebased('=SUM(A:B)', [0, 2], [5, 3])).toBe('=SUM(B:C)');
    expect(rebased('=SUM($A:B)', [0, 2], [5, 3])).toBe('=SUM($A:C)');
    expect(rebased('=SUM(2:3)', [0, 2], [5, 3])).toBe('=SUM(7:8)');
    expect(rebased('=SUM(2:$3)', [0, 2], [5, 3])).toBe('=SUM(7:$3)');
  });

  it('moves positional references too, and keeps their sheets', () => {
    expect(rebased('=SUM(Sheet1:Sheet3!A1)', [0, 2], [3, 3])).toBe('=SUM(Sheet1:Sheet3!B4)');
    expect(rebased('=SUM(Sheet1:Sheet3!A1:$B2)', [0, 2], [3, 3])).toBe('=SUM(Sheet1:Sheet3!B4:$B5)');
    expect(rebased('=SUM(Ghost!A1)', [0, 2], [3, 3])).toBe('=SUM(Ghost!B4)');
  });

  it('copes with a node that holds no references', () => {
    expect(rebased('=SUM({1,2;3,4})+50%', [0, 2], [3, 3])).toBe('=SUM({1,2;3,4})+50%');
  });
});

describe('dependencies and volatility', () => {
  it('lists a whole-column range as one axis rectangle and never as cells', () => {
    const { stable, current } = setup();
    const deps = collectStableDependencies(stable('=SUM(B:C)+SUM(2:3)+A1'));
    expect(deps.cells.size).toBe(1);
    expect(deps.ranges).toEqual([
      { startKey: `*:${current.getColId(1)}`, endKey: `*:${current.getColId(2)}`, axis: 'column' },
      { startKey: `${current.getRowId(1)}:*`, endKey: `${current.getRowId(2)}:*`, axis: 'row' },
    ]);
  });

  it('lists a plain range without an axis', () => {
    const { stable } = setup();
    const [range] = collectStableDependencies(stable('=SUM(A1:B2)')).ranges;
    expect(range).not.toHaveProperty('axis');
  });

  it('does not list a 3D reference', () => {
    const { stable } = setup();
    const deps = collectStableDependencies(stable('=SUM(Sheet1:Sheet3!A1)+SUM(Sheet1:Sheet3!A1:B2)'));
    expect(deps.cells.size).toBe(0);
    expect(deps.ranges).toHaveLength(0);
  });

  it('does not treat the shape of a plain reference as a dependency, but does for a spill reference', () => {
    const { stable } = setup();
    expect(collectStableDependencies(stable('=ROWS(A:A)+ROW(1:1)+COLUMNS(A1:B2)')).ranges).toHaveLength(0);
    expect(collectStableDependencies(stable('=ROWS(A1#)')).cells.size).toBe(1);
    expect(collectStableDependencies(stable('=ROWS(A1)')).cells.size).toBe(0);
  });

  it('flags 3D references and the range operator over computed references as volatile', () => {
    const { stable } = setup();
    expect(formulaIsVolatile(stable('=SUM(Sheet1:Sheet3!A1)'))).toBe(true);
    expect(formulaIsVolatile(stable('=1+SUM(A1:INDEX(A1:A5,3))'))).toBe(true);
    expect(formulaIsVolatile(stable('=SUM(A1:OFFSET(A1,1,1))'))).toBe(true);
    expect(formulaIsVolatile(stable('=NOW()'))).toBe(true);
  });

  it('does not flag whole-column ranges, unions, intersections, @, # or arrays', () => {
    const { stable } = setup();
    for (const f of ['=SUM(A:A)', '=SUM(1:1)', '=SUM((A1,B1))', '=SUM(A1:B2 B1:C3)', '=@A1:A3', '=SUM(A1#)', '=SUM({1,2})', '=A1%', '=A1:B2']) {
      expect(formulaIsVolatile(stable(f))).toBe(false);
    }
  });
});

describe('hand-built and damaged nodes', () => {
  it('prints a node with parts missing as what is there', () => {
    const { text } = setup();
    expect(text({ type: 'cell' })).toBe('=#REF!');
    expect(text({ type: 'operator' })).toBe('=');
    expect(text({ type: 'operator', operator: '+', left: { type: 'number', value: 1 } })).toBe('=1+');
    expect(text({ type: 'operator', operator: '+', right: { type: 'number', value: 1 } })).toBe('=+1');
    expect(text({ type: 'operator', operator: '??', left: { type: 'number', value: 1 }, right: { type: 'number', value: 2 } })).toBe('=1??2');
    expect(text({ type: 'variable', value: 'Total' })).toBe('=Total');
  });

  it('copies a formula whose named sheet has gone from the sheet list', () => {
    const { stable, text, resolver, current, sheets } = setup();
    const node = stable('=Sheet2!A1+1');
    sheets.delete('Sheet2');
    // The ids still belong to the removed sheet's order maps, so the copy keeps them untouched.
    expect(text(adjustStableAstForCopy(node, resolver, current, 0, 0, 1, 1))).toBe('=#REF!+1');
  });

  it('finds a volatile function on either side of an operator', () => {
    const { stable } = setup();
    expect(formulaIsVolatile(stable('=NOW()+1'))).toBe(true);
    expect(formulaIsVolatile(stable('=1+NOW()'))).toBe(true);
    expect(formulaIsVolatile(stable('=1+2'))).toBe(false);
  });
});
