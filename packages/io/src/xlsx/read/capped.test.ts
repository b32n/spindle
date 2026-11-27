import { abbreviate, CappedWarnings, flat, MAX_ITEM_WARNINGS, MAX_QUOTED_CHARS, truncate } from './capped';

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

describe('abbreviate', () => {
  it('leaves text of up to the limit as it is', () => {
    expect(abbreviate('')).toBe('');
    expect(abbreviate('a'.repeat(MAX_QUOTED_CHARS))).toBe('a'.repeat(MAX_QUOTED_CHARS));
    expect(abbreviate('abcdef', 6)).toBe('abcdef');
  });

  it('cuts longer text to the limit with an ellipsis as the last character', () => {
    expect(abbreviate('abcdefg', 6)).toBe('abcde…');
    const out = abbreviate('x'.repeat(10_000_000));
    expect(out).toBe(`${'x'.repeat(MAX_QUOTED_CHARS - 1)}…`);
    expect(out).toHaveLength(MAX_QUOTED_CHARS);
  });

  it('never leaves half a surrogate pair', () => {
    const pairs = '\u{1F600}'.repeat(100);
    for (const max of [6, 7, 8, 9]) {
      const out = abbreviate(pairs, max);
      expect(out.endsWith('…')).toBe(true);
      expect(out.length).toBeLessThanOrEqual(max);
      expect(out.slice(0, -1)).toBe('\u{1F600}'.repeat((out.length - 1) / 2));
    }
  });
});

describe('truncate', () => {
  it('keeps text of up to the limit and cuts longer text to the limit', () => {
    expect(truncate('abc', 3)).toBe('abc');
    expect(truncate('abcdef', 3)).toBe('abc');
    expect(truncate('x'.repeat(5_000_000), 4)).toBe('xxxx');
  });
  it('does not leave half a surrogate pair', () => {
    expect(truncate('\u{1F600}\u{1F600}', 3)).toBe('\u{1F600}');
    expect(truncate('\u{1F600}\u{1F600}', 2)).toBe('\u{1F600}');
    expect(truncate('a\u{1F600}', 2)).toBe('a');
  });
});

describe('flat', () => {
  it('is the same text: short text as it is, long text copied, whatever it holds (two-byte characters, surrogate pairs, NUL)', () => {
    for (const text of ['', 'a', 'twelve chars', 'thirteen char', `x\0y${'z'.repeat(100)}`, '\u20ac'.repeat(1000), '\u{1F600}'.repeat(50), 'A1:XFD1048576'])
      expect(flat(text)).toBe(text);
    const big = 'q'.repeat(5_000_000);
    expect(flat(big.slice(10, 4_000_010))).toHaveLength(4_000_000);
    expect(flat(`${'a'.repeat(20)}${'b'.repeat(20)}`)).toBe(`${'a'.repeat(20)}${'b'.repeat(20)}`);
  });

  it('keeps its meaning once the engine has compiled it (a string cut from a long one, a thousand times)', () => {
    const source = 'abcdefghijklmnopqrstuvwxyz'.repeat(100);
    for (let i = 0; i < 20_000; i++) {
      const piece = source.slice(i % 50, (i % 50) + 13 + (i % 7));
      if (flat(piece) !== piece) throw new Error(`flat changed ${piece}`);
    }
  });

  it('is applied by truncate and abbreviate, whatever their length', () => {
    expect(abbreviate('x'.repeat(40))).toBe('x'.repeat(40));
    expect(truncate('y'.repeat(40), 40)).toBe('y'.repeat(40));
  });
});
