import { DEFAULT_IMPORT_LIMITS } from '../../common/safe-zip';
import { LimitLoss, type CutKind } from './limit-loss';

const ALL: CutKind[] = ['cells', 'scan', 'sheets', 'rows', 'ranges', 'columns', 'extLst', 'formulas', 'parts', 'sharedStrings', 'relationships', 'definedNames', 'sheetNames', 'elements', 'numberFormats'];

describe('LimitLoss', () => {
  it('has no items when nothing was cut', () => {
    const loss = new LimitLoss(DEFAULT_IMPORT_LIMITS);
    expect(loss.items()).toEqual([]);
    expect(loss.has('cells')).toBe(false);
  });

  it('counts what is cut by kind, lists the first places (without repeats, quoted shortened) and says whether it is the first of its kind', () => {
    const loss = new LimitLoss(DEFAULT_IMPORT_LIMITS);
    expect(loss.add('rows', 'S1', 5)).toBe(true);
    expect(loss.add('rows', 'S1', 2)).toBe(false);
    for (const sheet of ['S2', 'S3', 'S4', 'S5', 'S6', 'S7']) loss.add('rows', sheet);
    loss.add('rows', 'x'.repeat(500));
    expect(loss.has('rows')).toBe(true);
    expect(loss.items()).toEqual([
      {
        id: 'xlsx.package.sizeLimit',
        label: 'row(s) with settings left out',
        count: 5 + 2 + 6 + 1,
        where: ['S1', 'S2', 'S3', 'S4', 'S5'],
        effect: `approximated: the cells are kept, the height, hidden state and style of the rows past maxRows (${DEFAULT_IMPORT_LIMITS.maxRows}) are not; raise it with the \`limits\` option of the import`,
        state: 'approximated',
      },
    ]);
  });

  it('leaves `where` out for a cut that has no place, and counts one by default', () => {
    const loss = new LimitLoss(DEFAULT_IMPORT_LIMITS);
    loss.add('numberFormats');
    loss.add('numberFormats');
    expect(loss.items()).toEqual([
      { id: 'xlsx.package.sizeLimit', label: 'number format code(s) over 255 characters', count: 2, effect: 'approximated: kept as written, but not read as a date or time (Excel allows 255 characters)', state: 'approximated' },
    ]);
  });

  it('names the limit and the option that raises it in every kind that is a limit, with the value of the limits it was made with, and starts with lost or approximated', () => {
    const limits = { ...DEFAULT_IMPORT_LIMITS, maxTotalCells: 111, maxScannedElements: 222, maxSheets: 3, maxRows: 4, maxRanges: 5, maxColumns: 6, maxExtensionChars: 7, maxSheetFormulaChars: 8, maxFormulaChars: 9, maxPartElements: 10, maxRelationships: 11 };
    const loss = new LimitLoss(limits);
    for (const kind of ALL) loss.add(kind, 'there');
    const items = loss.items();
    expect(items.map((i) => i.label)).toHaveLength(ALL.length);
    for (const item of items) {
      expect(item.id).toBe('xlsx.package.sizeLimit');
      expect(item.effect).toMatch(/^(lost|approximated): /);
      expect(item.state === 'approximated').toBe(item.effect.startsWith('approximated'));
    }
    const text = items.map((i) => i.effect).join('\n');
    for (const value of ['111', '222', '(3)', '(4)', '(5)', '(6)', '(7)', '(8)', '(9)', '(10)', '(11)']) expect(text).toContain(value);
    for (const option of ['maxTotalCells', 'maxScannedElements', 'maxSheets', 'maxRows', 'maxRanges', 'maxColumns', 'maxExtensionChars', 'maxSheetFormulaChars', 'maxFormulaChars', 'maxPartElements', 'maxRelationships']) expect(text).toContain(option);
    expect(items.filter((i) => /limits` option/.test(i.effect))).toHaveLength(ALL.length - 3);
  });
});
