import type { PptxParagraph, PptxRun, PptxTextBody } from '@b32nio/spindle-shared/export/pptx';
import type { RichTextDoc } from '../text/model';
import {
  bodyPropsToPptx,
  bodyStyleFromPptx,
  levelStyleFromAttrs,
  newTextTally,
  paragraphAttrsFromPptx,
  textFromPptx,
  textToPptx,
} from './pptx-text';

const run = (text: string, extra: Partial<PptxRun> = {}): PptxRun => ({ kind: 'text', text, ...extra });
const para = (runs: PptxRun[], style: PptxParagraph['style'] = {}, level = 0): PptxParagraph => ({ level, style, runs });
const body = (paragraphs: PptxParagraph[], props: PptxTextBody['props'] = {}): PptxTextBody => ({ props, paragraphs });

describe('body properties', () => {
  it('defaults the insets and merges the ones the file states', () => {
    const t = newTextTally();
    expect(bodyStyleFromPptx({}, t).insets).toEqual({ l: 9.6, t: 4.8, r: 9.6, b: 4.8 });
    expect(bodyStyleFromPptx({ insets: { l: 0, t: undefined } }, t).insets).toEqual({ l: 0, t: 4.8, r: 9.6, b: 4.8 });
    expect(bodyStyleFromPptx({}, t).vAlign).toBe('top');
    expect(bodyStyleFromPptx({}, t, 'middle').vAlign).toBe('middle');
    expect(bodyStyleFromPptx({ anchor: 'bottom' }, t, 'middle').vAlign).toBe('bottom');
  });
  it('maps the switches, counting vertical directions it cannot draw', () => {
    const t = newTextTally();
    const s = bodyStyleFromPptx(
      { anchorCtr: true, wrap: 'none', vert: 'eaVert', autofit: { kind: 'shrink', fontScale: 0.5, lnSpcReduction: 0.2 }, columns: 2, columnSpacing: 12 },
      t
    );
    expect(s).toMatchObject({ anchorCtr: true, wrap: false, vert: 'eaVert', autofit: { mode: 'shrink', fontScale: 0.5, lnSpcReduction: 0.2 }, columns: { count: 2, gap: 12 } });
    expect(bodyStyleFromPptx({ vert: 'horz', wrap: 'square', autofit: { kind: 'none' }, columns: 1 }, t)).toMatchObject({ autofit: { mode: 'none' } });
    expect(bodyStyleFromPptx({ columns: 3 }, t).columns).toEqual({ count: 3 });
    expect(bodyStyleFromPptx({ vert: 'vert270' }, t).vert).toBeUndefined();
    expect(t.verticalText).toBe(1);
  });
  it('writes body properties back', () => {
    expect(bodyPropsToPptx(undefined)).toEqual({});
    expect(bodyPropsToPptx({ padding: 8 }).insets).toEqual({ l: 8, t: 8, r: 8, b: 8 });
    expect(bodyPropsToPptx({ padding: 8, insets: { l: 1, t: 2, r: 3, b: 4 } }).insets).toEqual({ l: 1, t: 2, r: 3, b: 4 });
    expect(
      bodyPropsToPptx({ vAlign: 'middle', anchorCtr: false, vert: 'vert', wrap: false, autofit: { mode: 'shrink', fontScale: 0.9, lnSpcReduction: 0.1 }, columns: { count: 2, gap: 5 } })
    ).toEqual({ anchor: 'middle', anchorCtr: false, vert: 'vert', wrap: 'none', autofit: { kind: 'shrink', fontScale: 0.9, lnSpcReduction: 0.1 }, columns: 2, columnSpacing: 5 });
    expect(bodyPropsToPptx({ autofit: { mode: 'resize' }, columns: { count: 2 }, wrap: true })).toEqual({ autofit: { kind: 'resize' }, columns: 2 });
    expect(bodyPropsToPptx({ columns: { count: 1 } })).toEqual({});
  });
});

describe('paragraph attributes', () => {
  it('maps everything a level style states', () => {
    const a = paragraphAttrsFromPptx(2, {
      align: 'justify',
      lineSpacingPts: 12,
      spaceBeforePct: 0.2,
      spaceAfter: 6,
      bullet: { kind: 'auto', scheme: 'romanUcPeriod', startAt: 3 },
      bulletColor: { kind: 'rgb', hex: '#00FF00' },
      bulletSizePct: 0.8,
      marginLeft: 20,
      indent: -10,
      tabStops: [{ pos: 5, align: 'ctr' }, { pos: 9, align: 'weird' }],
    });
    expect(a).toEqual({
      align: 'justify',
      indent: 2,
      lineHeightPx: 16,
      spaceBeforePct: 20,
      spaceAfter: 8,
      bullet: { kind: 'auto', scheme: 'romanUpper', startAt: 3, color: { kind: 'rgb', hex: '#00FF00' }, sizePct: 80 },
      marginLeft: 20,
      firstLineIndent: -10,
      tabStops: [{ pos: 5, align: 'ctr' }],
    });
    expect(paragraphAttrsFromPptx(0, { lineHeight: 1.5, spaceBefore: 3, spaceAfterPct: 0.1, bullet: { kind: 'none' } })).toEqual({
      lineHeight: 1.5,
      spaceBefore: 4,
      spaceAfterPct: 10,
      bullet: { kind: 'none' },
    });
    expect(paragraphAttrsFromPptx(0, { bullet: { kind: 'char', char: '-' } }).bullet).toEqual({ kind: 'char', char: '-' });
    expect(paragraphAttrsFromPptx(0, { bullet: { kind: 'auto', scheme: 'mystery' } }).bullet).toEqual({ kind: 'auto', scheme: 'arabic' });
    expect(paragraphAttrsFromPptx(0, {})).toEqual({});
  });

  it('writes attributes back to a level style over the body defaults', () => {
    const d = { fontSize: 18 };
    expect(levelStyleFromAttrs(undefined, d)).toEqual(d);
    expect(levelStyleFromAttrs({ align: 'right', lineHeightPx: 24, spaceBefore: 8, spaceAfterPct: 50 }, d)).toEqual({
      fontSize: 18, align: 'right', lineSpacingPts: 18, spaceBefore: 6, spaceAfterPct: 0.5,
    });
    expect(levelStyleFromAttrs({ lineHeight: 1.3, spaceBeforePct: 0, spaceAfter: 4 }, {})).toEqual({ lineHeight: 1.3, spaceBeforePct: 0, spaceAfter: 3 });
    expect(levelStyleFromAttrs({ lineHeightPx: -1, lineHeight: NaN, spaceBeforePct: -5, spaceAfterPct: -5 }, {})).toEqual({});
    expect(
      levelStyleFromAttrs({ bullet: { kind: 'auto', scheme: 'alphaLower', startAt: 2, color: { kind: 'rgb', hex: '#111111' }, sizePct: 90 }, marginLeft: 30, firstLineIndent: -15, tabStops: [{ pos: 4, align: 'r' }, { pos: NaN, align: 'l' }] }, {})
    ).toEqual({
      bullet: { kind: 'auto', scheme: 'alphaLcPeriod', startAt: 2 },
      bulletColor: { kind: 'rgb', hex: '#111111' },
      bulletSizePct: 0.9,
      marginLeft: 30,
      indent: -15,
      tabStops: [{ pos: 4, align: 'r' }],
    });
    expect(levelStyleFromAttrs({ bullet: { kind: 'char' } }, {}).bullet).toEqual({ kind: 'char', char: '•' });
    expect(levelStyleFromAttrs({ bullet: { kind: 'auto' } }, {}).bullet).toEqual({ kind: 'auto', scheme: 'arabicPeriod' });
    expect(levelStyleFromAttrs({ bullet: { kind: 'none' } }, {}).bullet).toEqual({ kind: 'none' });
    expect(levelStyleFromAttrs({ bullet: { kind: 'char', char: 'x', sizePct: NaN } }, {})).not.toHaveProperty('bulletSizePct');
    // the older list types and the legacy 28 px step per level
    expect(levelStyleFromAttrs({ listType: 'bullet', indent: 1 }, {})).toMatchObject({ bullet: { kind: 'char', char: '•' }, marginLeft: 56, indent: -28 });
    expect(levelStyleFromAttrs({ listType: 'number', firstLineIndent: -5 }, {})).toMatchObject({ bullet: { kind: 'auto', scheme: 'arabicPeriod' }, marginLeft: 28, indent: -5 });
    expect(levelStyleFromAttrs({ indent: 2 }, {})).toEqual({ marginLeft: 56 });
    expect(levelStyleFromAttrs({ firstLineIndent: 7 }, {})).toEqual({ indent: 7 });
    expect(levelStyleFromAttrs({ marginLeft: 10 }, {})).toEqual({ marginLeft: 10 });
    expect(levelStyleFromAttrs({ bullet: null, listType: 'none' }, {})).toEqual({});
  });
});

describe('text from the interchange', () => {
  it('puts a shared size, colour and family on the body and marks only what differs', () => {
    const red = { kind: 'rgb' as const, hex: '#FF0000' };
    const { richText, bodyStyle } = textFromPptx(
      body([
        para([run('plain'), run('big', { size: 30 }), run('red', { color: red }), run('mono', { font: 'Courier' })], { fontSize: 18, color: { kind: 'theme', slot: 'dk1' }, font: 'minor' }),
      ]),
      newTextTally()
    );
    expect(bodyStyle).toMatchObject({ fontSize: 24, fontFamily: 'minor', color: { kind: 'theme', slot: 'dk1' } });
    const marks = richText.content[0].content!.map((i) => (i.type === 'text' ? (i.marks ?? []).map((m) => m.type) : []));
    expect(marks).toEqual([[], ['fontSize'], ['textColor'], ['fontFamily']]);
  });

  it('marks bold only when it differs from a shared bold, and italic, underline, strike, link, highlight, baseline, caps, spacing, lang', () => {
    const t = newTextTally();
    const all = textFromPptx(body([para([run('a', { bold: true }), run('b', { bold: true })])]), t);
    expect(all.bodyStyle.bold).toBe(true);
    expect(all.richText.content[0].content!.every((i) => i.type === 'text' && !i.marks)).toBe(true);
    const mixed = textFromPptx(body([para([run('a', { bold: true }), run('b')], { bold: false })]), t);
    expect(mixed.bodyStyle.bold).toBeUndefined();
    const styled = textFromPptx(
      body([
        para([
          run('x', {
            italic: true, underline: 'sng', strike: 'double', link: { kind: 'url', href: 'https://a.example' },
            highlight: { kind: 'rgb', hex: '#FFFF00' }, baseline: 0.3, caps: 'small', spacing: 3, lang: 'en-US',
          }),
          run('y', { link: { kind: 'action', target: 'ppaction://x' }, baseline: 0, spacing: 0 }),
        ]),
      ]),
      t
    );
    const [x, y] = styled.richText.content[0].content!;
    expect(x).toMatchObject({ marks: [
      { type: 'italic' }, { type: 'underline' }, { type: 'strikethrough' }, { type: 'link', attrs: { href: 'https://a.example' } },
      { type: 'highlight' }, { type: 'baseline', attrs: { pct: 30 } }, { type: 'caps', attrs: { kind: 'small' } }, { type: 'charSpacing', attrs: { px: 4 } }, { type: 'lang' },
    ] });
    expect(y).not.toHaveProperty('marks');
  });

  it('keeps slide number and date fields, other fields as text, and splits a paragraph at a line break', () => {
    const t = newTextTally();
    const { richText } = textFromPptx(
      body([
        para(
          [
            run('1', { kind: 'field', fieldType: 'slidenum' } as Partial<PptxRun>),
            { kind: 'field', fieldType: 'datetime1', text: '1/1' },
            { kind: 'field', fieldType: 'datetime12', text: '10:00' },
            { kind: 'field', fieldType: 'datetime', text: 'now' },
            { kind: 'field', fieldType: 'other', text: 'o' },
            { kind: 'field', text: 'untyped' },
            { kind: 'break', text: '' },
            run(''),
            run('after'),
          ],
          { bullet: { kind: 'char', char: '*' }, spaceBefore: 5, marginLeft: 20, indent: -10 }
        ),
        para([]),
      ]),
      t
    );
    const [first, second, empty] = richText.content;
    expect(first.content!.map((i) => (i.type === 'field' ? i.attrs : 'text'))).toEqual([
      { kind: 'slidenum' }, { kind: 'datetime', format: 'date' }, { kind: 'datetime', format: 'time' }, { kind: 'datetime', format: 'date' }, 'text', 'text',
    ]);
    expect(second.attrs).toMatchObject({ bullet: { kind: 'none' }, spaceBefore: 0, spaceBeforePct: null, firstLineIndent: 0, marginLeft: 20 });
    expect(empty).toEqual({ type: 'paragraph' });
    expect(t.lineBreaks).toBe(1);
    // no paragraphs at all still gives a valid document
    expect(textFromPptx(body([]), t).richText.content).toEqual([{ type: 'paragraph' }]);
    // a break in a paragraph without a first-line indent does not invent one
    const noIndent = textFromPptx(body([para([run('a'), { kind: 'break', text: '' }, run('b')], { align: 'center' })]), t);
    expect(noIndent.richText.content[1].attrs).not.toHaveProperty('firstLineIndent');
  });
});

describe('text to the interchange', () => {
  const doc: RichTextDoc = {
    type: 'doc',
    content: [
      {
        type: 'paragraph',
        attrs: { indent: 1, align: 'center' },
        content: [
          { type: 'text', text: 'a', marks: [
            { type: 'bold' }, { type: 'italic' }, { type: 'underline' }, { type: 'strikethrough' },
            { type: 'link', attrs: { href: 'https://x.example' } },
            { type: 'textColor', attrs: { color: { kind: 'rgb', hex: '#010203' } } },
            { type: 'fontFamily', attrs: { family: 'Arial' } },
            { type: 'fontSize', attrs: { size: 24 } },
            { type: 'highlight', attrs: { color: { kind: 'rgb', hex: '#FFFF00' } } },
            { type: 'baseline', attrs: { pct: -30 } },
            { type: 'caps', attrs: { kind: 'all' } },
            { type: 'charSpacing', attrs: { px: 4 } },
            { type: 'lang', attrs: { lang: 'fr' } },
          ] },
          { type: 'text', text: '' },
          { type: 'text', text: 'plain' },
          { type: 'field', attrs: { kind: 'slidenum' }, content: [{ type: 'text', text: '3', marks: [{ type: 'bold' }] }], marks: [{ type: 'italic' }] },
          { type: 'field', attrs: { kind: 'datetime', format: 'time' }, content: [{ type: 'text', text: '10:00' }] },
          { type: 'field', attrs: { kind: 'datetime' } },
          { type: 'field', attrs: { kind: 'footer' }, content: [{ type: 'text', text: 'Foot' }] },
          { type: 'field', attrs: { kind: 'datetime', format: 'bogus' as never } },
          { type: 'text', text: 'c', marks: [{ type: 'textColor', attrs: { color: { kind: 'theme', slot: 'accent2' } } }] },
        ],
      },
      { type: 'paragraph' },
    ],
  };

  it('writes runs with their marks, fields with their types and an empty paragraph with the default size', () => {
    const out = textToPptx(doc, { fontSize: 24, bold: true, color: { kind: 'rgb', hex: '#000000' }, fontFamily: 'minor', vAlign: 'top' });
    const [p, empty] = out.paragraphs;
    expect(p.level).toBe(1);
    expect(p.runs[0]).toEqual({
      kind: 'text', text: 'a', bold: true, italic: true, underline: 'sng', strike: 'single', link: { kind: 'url', href: 'https://x.example' },
      color: { kind: 'rgb', hex: '#010203' }, font: 'Arial', size: 18, highlight: { kind: 'rgb', hex: '#FFFF00' }, baseline: -0.3, caps: 'all', spacing: 3, lang: 'fr',
    });
    expect(p.runs.slice(1).map((r) => [r.kind, r.fieldType, r.text])).toEqual([
      ['text', undefined, 'plain'], ['field', 'slidenum', '3'], ['field', 'datetime11', '10:00'], ['field', 'datetime1', ''], ['text', undefined, 'Foot'], ['field', 'datetime1', ''], ['text', undefined, 'c'],
    ]);
    expect(p.runs[2]).toMatchObject({ bold: true, italic: true });
    expect(p.style).toMatchObject({ fontSize: 18, bold: true, color: { kind: 'rgb', hex: '#000000' }, font: 'minor', align: 'center', marginLeft: 28 });
    expect(empty).toEqual({ level: 0, style: p.style && empty.style, runs: [], endSize: 18 });
    expect(out.props).toMatchObject({ anchor: 'top' });
  });

  it('ignores marks with unusable values and clamps the level', () => {
    const bad: RichTextDoc = {
      type: 'doc',
      content: [{ type: 'paragraph', attrs: { indent: 99 }, content: [{ type: 'text', text: 'z', marks: [
        { type: 'link', attrs: { href: '' } }, { type: 'textColor', attrs: { color: 'red' } }, { type: 'textColor' },
        { type: 'fontFamily', attrs: { family: 3 } }, { type: 'fontSize', attrs: { size: -2 } }, { type: 'highlight', attrs: { color: null } },
        { type: 'baseline', attrs: { pct: 0 } }, { type: 'caps', attrs: { kind: 'x' } }, { type: 'charSpacing', attrs: { px: 'a' } },
        { type: 'lang', attrs: { lang: '' } }, { type: 'unknown' },
      ] }] }],
    };
    const out = textToPptx(bad, undefined);
    expect(out.paragraphs[0].level).toBe(8);
    expect(out.paragraphs[0].runs[0]).toEqual({ kind: 'text', text: 'z' });
    expect(out.props).toEqual({});
    expect(out.paragraphs[0].style).toEqual({ marginLeft: 224 });
    expect(textToPptx({ type: 'doc', content: [{ type: 'paragraph' }] }, { fontSize: 0, bold: false, color: { kind: 'x' } as never, fontFamily: '' }).paragraphs[0]).toEqual({ level: 0, style: { bold: false }, runs: [] });
  });
});
