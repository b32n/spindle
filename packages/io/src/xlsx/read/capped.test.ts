import { CappedWarnings, MAX_ITEM_WARNINGS } from './capped';

describe('CappedWarnings', () => {
  it('passes the first warnings as they are and counts the rest in one closing message', () => {
    const out: string[] = [];
    const w = new CappedWarnings((m) => out.push(m), 'link warnings', 3);
    for (let i = 0; i < 10; i++) w.add(`w${i}`);
    w.close();
    expect(out).toEqual(['w0', 'w1', 'w2', '7 more link warnings were left out.']);
  });

  it('says nothing more when the cap was not passed (exactly at it too)', () => {
    const out: string[] = [];
    const w = new CappedWarnings((m) => out.push(m), 'x', 2);
    w.add('a');
    w.add('b');
    w.close();
    expect(out).toEqual(['a', 'b']);
  });

  it('has a default cap of MAX_ITEM_WARNINGS', () => {
    const out: string[] = [];
    const w = new CappedWarnings((m) => out.push(m), 'things');
    for (let i = 0; i < 1_000_000; i++) w.add('same');
    w.close();
    expect(out).toHaveLength(MAX_ITEM_WARNINGS + 1);
    expect(out[MAX_ITEM_WARNINGS]).toBe(`${1_000_000 - MAX_ITEM_WARNINGS} more things were left out.`);
  });
});
