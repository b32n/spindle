import { slidesSchema } from './schema';
import type { ParagraphAttrs, RichTextDoc } from './model';
import {
  DEFAULT_BULLET_CHAR,
  flattenLevelStyles,
  formatOrdinal,
  normalizeTabStops,
  resolveParagraph,
  safeBulletChar,
} from './paragraph';

describe('resolveParagraph', () => {
  it('keeps the legacy layout: 28px per level, a marker gutter, 1.2 line height', () => {
    expect(resolveParagraph({})).toMatchObject({ marker: null, paddingLeft: 0, textIndent: 0, lineHeight: { mult: 1.2 }, spaceBefore: null });
    expect(resolveParagraph({ indent: 2 }).paddingLeft).toBe(56);
    const bullet = resolveParagraph({ listType: 'bullet', indent: 1 });
    expect(bullet.marker).toMatchObject({ kind: 'bullet', char: DEFAULT_BULLET_CHAR, sizePct: 100 });
    expect(bullet.paddingLeft).toBe(56);
    expect(bullet.markerLeft).toBe(8);
    expect(resolveParagraph({ listType: 'number' }).markerLeft).toBe(2);
    expect(resolveParagraph({ listType: 'none' }).marker).toBeNull();
  });

  it('a bullet attr overrides listType', () => {
    expect(resolveParagraph({ listType: 'bullet', bullet: { kind: 'none' } }).marker).toBeNull();
    const m = resolveParagraph({ listType: 'bullet', bullet: { kind: 'auto', scheme: 'romanUpper', startAt: 4.9, sizePct: 150, color: { kind: 'theme', slot: 'accent1' } } }).marker;
    expect(m).toEqual({ kind: 'number', char: '•', scheme: 'romanUpper', startAt: 4, color: { kind: 'theme', slot: 'accent1' }, sizePct: 150 });
    const c = resolveParagraph({ bullet: { kind: 'char', char: '➤' } }).marker;
    expect(c).toMatchObject({ kind: 'bullet', char: '➤', scheme: 'arabic', startAt: null, color: null });
  });

  it('treats a crafted bullet as absent and clamps sizes', () => {
    expect(resolveParagraph({ bullet: { kind: 'evil' } as never }).marker).toBeNull();
    expect(resolveParagraph({ bullet: null }).marker).toBeNull();
    expect(resolveParagraph({ bullet: 'x' as never }).marker).toBeNull();
    const m = resolveParagraph({ bullet: { kind: 'auto', scheme: 'x' as never, sizePct: 9999, startAt: NaN, color: 'red' as never } }).marker;
    expect(m).toMatchObject({ scheme: 'arabic', sizePct: 400, startAt: null, color: null });
    expect(resolveParagraph({ bullet: { kind: 'char', sizePct: -5 } }).marker?.sizePct).toBe(100);
  });

  it('sanitises the glyph', () => {
    expect(safeBulletChar('★')).toBe('★');
    expect(safeBulletChar('abcdefgh')).toBe('abcd');
    for (const bad of [undefined, 5, '', '  ', 'a\u0000', '\n']) expect(safeBulletChar(bad)).toBe(DEFAULT_BULLET_CHAR);
  });

  it('resolves line height forms, space forms and the margin / indent pair', () => {
    expect(resolveParagraph({ lineHeightPx: 24, lineHeight: 2 }).lineHeight).toEqual({ px: 24 });
    expect(resolveParagraph({ lineHeightPx: 0, lineHeight: 2 }).lineHeight).toEqual({ mult: 2 });
    expect(resolveParagraph({ lineHeight: '9' as never }).lineHeight).toEqual({ mult: 9 });
    expect(resolveParagraph({ lineHeight: 'x' as never }).lineHeight).toEqual({ mult: 1.2 });
    expect(resolveParagraph({ spaceBeforePct: 50, spaceBefore: 9, spaceAfter: 6 })).toMatchObject({ spaceBefore: { em: 0.5 }, spaceAfter: { px: 6 } });
    expect(resolveParagraph({ spaceBeforePct: 0, spaceBefore: 9 }).spaceBefore).toEqual({ px: 9 });
    // Hanging bullet: block at 40px, marker pulled 30px back.
    const hang = resolveParagraph({ bullet: { kind: 'char' }, marginLeft: 40, firstLineIndent: -30 });
    expect(hang).toMatchObject({ paddingLeft: 40, markerLeft: 10, textIndent: 0 });
    expect(resolveParagraph({ bullet: { kind: 'char' }, marginLeft: 10, firstLineIndent: -30 }).markerLeft).toBe(0);
    // Without a marker the first-line offset indents the text.
    expect(resolveParagraph({ marginLeft: 20, firstLineIndent: 12 })).toMatchObject({ paddingLeft: 20, textIndent: 12 });
    expect(resolveParagraph({ marginLeft: -5 }).paddingLeft).toBe(0);
  });
});

describe('formatOrdinal', () => {
  it('formats every scheme', () => {
    expect(formatOrdinal(3, 'arabic')).toBe('3');
    expect(formatOrdinal(1, 'alphaLower')).toBe('a');
    expect(formatOrdinal(26, 'alphaLower')).toBe('z');
    expect(formatOrdinal(27, 'alphaUpper')).toBe('AA');
    expect(formatOrdinal(52, 'alphaLower')).toBe('az');
    expect(formatOrdinal(4, 'romanLower')).toBe('iv');
    expect(formatOrdinal(1994, 'romanUpper')).toBe('MCMXCIV');
    expect(formatOrdinal(99999, 'romanLower')).toBe('mmmcmxcix');
    expect(formatOrdinal(0, 'romanLower')).toBe('0');
    expect(formatOrdinal(NaN, 'alphaLower')).toBe('1');
  });
});

describe('normalizeTabStops', () => {
  it('keeps valid stops sorted, defaults the alignment, drops bad entries', () => {
    expect(normalizeTabStops([{ pos: 90, align: 'r' }, { pos: 30, align: 'x' as never }, { pos: -1, align: 'l' }, { pos: NaN, align: 'l' }, null, { pos: 60, align: 'dec' }, { pos: 70, align: 'ctr' }])).toEqual([
      { pos: 30, align: 'l' },
      { pos: 60, align: 'dec' },
      { pos: 70, align: 'ctr' },
      { pos: 90, align: 'r' },
    ]);
    expect(normalizeTabStops(undefined)).toEqual([]);
    expect(normalizeTabStops('x')).toEqual([]);
  });
});

describe('flattenLevelStyles', () => {
  const doc: RichTextDoc = {
    type: 'doc',
    content: [{ type: 'paragraph' }, { type: 'paragraph', attrs: { indent: 1, marginLeft: 5 } }, { type: 'paragraph', attrs: { indent: 20 } }, { type: 'paragraph', attrs: { indent: -3 } }],
  };
  it('lets a paragraph keep what it sets and inherit the rest', () => {
    const levels = { 0: { bullet: { kind: 'char' as const }, marginLeft: 10 }, 1: { marginLeft: 99, spaceAfter: 4 } };
    const out = flattenLevelStyles(doc, levels);
    expect(out.content[0].attrs).toEqual({ bullet: { kind: 'char' }, marginLeft: 10 });
    expect(out.content[1].attrs).toEqual({ marginLeft: 5, spaceAfter: 4, indent: 1 });
    // Levels clamp to 0..8: 20 -> 8 (no style), -3 -> 0.
    expect(out.content[2]).toBe(doc.content[2]);
    expect(out.content[3].attrs).toMatchObject({ bullet: { kind: 'char' }, indent: -3 });
  });
  it('is the identity without level styles', () => {
    expect(flattenLevelStyles(doc, undefined)).toBe(doc);
  });
});

describe('schema toDOM for K3b attrs', () => {
  const dom = (attrs: Record<string, unknown>) => {
    const node = slidesSchema.nodes.paragraph.create(attrs);
    const spec = slidesSchema.nodes.paragraph.spec.toDOM!(node) as [string, Record<string, string>, number];
    return spec[1];
  };
  it('emits no K3b styling for legacy paragraphs', () => {
    const a = dom({});
    expect(a.style).toBe('text-align:left;position:relative;line-height:calc(1.2 * var(--spindle-ls, 1));');
    expect(a['data-list']).toBe('none');
    expect(a['data-bullet-char']).toBeUndefined();
  });
  it('renders bullet, spacing, margin and indent', () => {
    const a = dom({
      bullet: { kind: 'char', char: '✓', sizePct: 80, color: { kind: 'rgb', hex: '#ff0000' } },
      lineHeightPx: 30,
      spaceBeforePct: 20,
      spaceAfter: 6,
      marginLeft: 36,
      firstLineIndent: -18,
    });
    expect(a['data-list']).toBe('bullet');
    expect(a['data-bullet-char']).toBe('✓');
    expect(a.style).toContain('line-height:calc(30px * var(--spindle-ls, 1));');
    expect(a.style).toContain('margin-top:0.2em;margin-bottom:6px;padding-left:36px;');
    expect(a.style).toContain('--pm-marker-left:18px;--pm-marker-size:80%;--pm-marker-color:#ff0000;');
    const t = dom({ bullet: { kind: 'auto', scheme: 'alphaUpper', startAt: 3 }, textIndent: 1 });
    expect(t['data-list']).toBe('number');
    expect(t['data-scheme']).toBe('alphaUpper');
    expect(t.style).toContain('--pm-start:3;');
    expect(dom({ marginLeft: 20, firstLineIndent: 10 }).style).toContain('text-indent:10px;');
  });
  it('cannot be used to inject CSS', () => {
    const a = dom({
      bullet: { kind: 'char', char: '";}body{x:y', color: { kind: 'rgb', hex: 'red;background:url(//evil)' }, sizePct: '1;x:y' },
      marginLeft: '5px;background:red',
      lineHeightPx: 'x;y',
      spaceBeforePct: '1;z',
    });
    expect(a.style).not.toMatch(/evil|background|;x:y|;y|;z/);
    expect(a['data-bullet-char']).toBe('";}b');
    expect(a.style).not.toContain('--pm-marker-color');
  });
  it('round-trips the attrs through the schema JSON', () => {
    const attrs: ParagraphAttrs = { bullet: { kind: 'auto', startAt: 2, scheme: 'romanLower' }, tabStops: [{ pos: 40, align: 'dec' }], firstLineIndent: -12, marginLeft: 12, lineHeightPx: 20, spaceBeforePct: 10, spaceAfterPct: 15 };
    const json = slidesSchema.nodeFromJSON({ type: 'doc', content: [{ type: 'paragraph', attrs }] }).toJSON();
    expect(json.content[0].attrs).toMatchObject(attrs);
  });
});
