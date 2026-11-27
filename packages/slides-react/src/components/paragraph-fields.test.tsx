import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { getBuiltinTheme, type ParagraphAttrs, type RichTextDoc } from '@b32nio/spindle-slides-core';
import { StaticRichText } from './elements/StaticRichText';
import { ParagraphFields } from './ParagraphFields';

const theme = getBuiltinTheme();
const render = (content: RichTextDoc['content']) => renderToStaticMarkup(<StaticRichText doc={{ type: 'doc', content }} theme={theme} />);
const para = (attrs: ParagraphAttrs, text = 'x') => ({ type: 'paragraph' as const, attrs, content: [{ type: 'text' as const, text }] });

type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const find = (node: React.ReactNode, pred: (e: React.ReactElement<P>) => boolean, out: React.ReactElement<P>[] = []) => {
  React.Children.forEach(node, (c) => {
    if (!React.isValidElement<P>(c)) return;
    if (pred(c)) out.push(c);
    find(c.props.children, pred, out);
  });
  return out;
};
const byId = (id: string) => (e: React.ReactElement<P>) => e.props['data-testid'] === id;
const tree = (attrs: ParagraphAttrs | undefined, onChange = jest.fn()) => ({ onChange, el: ParagraphFields({ attrs, onChange }) });
const one = (el: React.ReactNode, id: string) => find(el, byId(id))[0];

describe('StaticRichText paragraph attributes', () => {
  it('keeps the legacy bullet and number markers', () => {
    const h = render([para({ listType: 'bullet' }), para({ listType: 'number' }, 'a'), para({ listType: 'number' }, 'b')]);
    expect(h).toContain('•');
    expect(h).toContain('>1.<');
    expect(h).toContain('>2.<');
    expect(h).toContain('padding-left:28px');
  });

  it('draws a custom glyph with colour and size, hanging at the margin', () => {
    const h = render([para({ bullet: { kind: 'char', char: '➤', sizePct: 80, color: { kind: 'rgb', hex: '#ff0000' } }, marginLeft: 40, firstLineIndent: -30 })]);
    expect(h).toContain('➤');
    expect(h).toContain('font-size:80%');
    expect(h).toContain('color:rgb(255, 0, 0)');
    expect(h).toContain('left:10px');
    expect(h).toContain('padding-left:40px');
  });

  it('numbers in each scheme, applying a start only where the list begins', () => {
    const attrs = (scheme: 'alphaUpper' | 'romanLower', startAt?: number): ParagraphAttrs => ({ bullet: { kind: 'auto', scheme, startAt } });
    const h = render([para(attrs('alphaUpper', 3)), para(attrs('alphaUpper', 3)), para({}), para(attrs('romanLower', 3))]);
    expect(h).toContain('>C.<');
    expect(h).toContain('>D.<');
    expect(h).toContain('>iii.<');
    expect(render([para(attrs('romanLower'))])).toContain('>i.<');
  });

  it('a "none" bullet overrides listType and a number run restarts after a plain paragraph', () => {
    expect(render([para({ listType: 'bullet', bullet: { kind: 'none' } })])).not.toContain('•');
    const h = render([para({ listType: 'number' }), para({}), para({ listType: 'number' }, 'z')]);
    expect(h.match(/>1\.</g)).toHaveLength(2);
  });

  it('renders exact line height, percent spacing and a text indent', () => {
    const h = render([para({ lineHeightPx: 30, spaceBeforePct: 50, spaceAfter: 6, marginLeft: 12, firstLineIndent: 8 })]);
    expect(h).toContain('line-height:30px');
    expect(h).toContain('margin-top:0.5em');
    expect(h).toContain('margin-bottom:6px');
    expect(h).toContain('text-indent:8px');
    const scaled = renderToStaticMarkup(<StaticRichText scaled doc={{ type: 'doc', content: [para({ lineHeightPx: 30 })] }} theme={theme} />);
    expect(scaled).toContain('calc(30px * var(--spindle-ls, 1))');
  });
});

describe('ParagraphFields', () => {
  it('shows defaults for a plain paragraph and hides bullet detail', () => {
    const { el } = tree(undefined);
    expect(one(el, 'para-bullet').props.value).toBe('default');
    expect(one(el, 'para-bullet-char')).toBeUndefined();
    expect(one(el, 'para-lh-mode').props.value).toBe('multiple');
    expect(one(el, 'para-lh').props.value).toBe(1.2);
    expect(one(el, 'para-margin-left').props.value).toBe(0);
    expect((el as React.ReactElement<P>).props['aria-label']).toBe('Paragraph options');
  });

  it('the bullet choice sets, clears or switches the bullet', () => {
    const t = tree({ bullet: { kind: 'char', char: '★' } });
    const sel = one(t.el, 'para-bullet');
    sel.props.onChange({ target: { value: 'default' } });
    sel.props.onChange({ target: { value: 'none' } });
    sel.props.onChange({ target: { value: 'auto' } });
    expect(t.onChange.mock.calls.map((c) => c[0])).toEqual([{ bullet: null }, { bullet: { kind: 'none' } }, { bullet: { kind: 'auto', char: '★' } }]);
    const fresh = tree(undefined);
    one(fresh.el, 'para-bullet').props.onChange({ target: { value: 'char' } });
    expect(fresh.onChange).toHaveBeenCalledWith({ bullet: { kind: 'char' } });
  });

  it('char bullets edit glyph, size and colour', () => {
    const t = tree({ bullet: { kind: 'char', char: '★', color: { kind: 'rgb', hex: '#00ff00' } } });
    expect(one(t.el, 'para-bullet-char').props.value).toBe('★');
    one(t.el, 'para-bullet-char').props.onChange({ target: { value: '✓' } });
    one(t.el, 'para-bullet-char').props.onChange({ target: { value: '' } });
    one(t.el, 'para-bullet-size').props.onValue(9999);
    one(t.el, 'para-bullet-size').props.onValue(1);
    expect(one(t.el, 'para-bullet-color').props.value).toBe('#00ff00');
    one(t.el, 'para-bullet-color').props.onChange({ target: { value: '#112233' } });
    const calls = t.onChange.mock.calls.map((c) => (c[0] as { bullet: Record<string, unknown> }).bullet);
    expect(calls[0]).toMatchObject({ kind: 'char', char: '✓' });
    expect(calls[1].char).toBeUndefined();
    expect(calls[2].sizePct).toBe(400);
    expect(calls[3].sizePct).toBe(25);
    expect(calls[4].color).toEqual({ kind: 'rgb', hex: '#112233' });
    expect(one(tree({ bullet: { kind: 'char' } }).el, 'para-bullet-char').props.value).toBe('•');
    expect(one(tree({ bullet: { kind: 'char', color: { kind: 'theme', slot: 'accent1' } } }).el, 'para-bullet-color').props.value).toBe('#000000');
    expect(one(tree({ bullet: { kind: 'char', color: { kind: 'rgb', hex: 'red' } } }).el, 'para-bullet-color').props.value).toBe('#000000');
  });

  it('numbered bullets edit scheme and start', () => {
    const t = tree({ bullet: { kind: 'auto', scheme: 'romanUpper', startAt: 4 } });
    expect(one(t.el, 'para-bullet-scheme').props.value).toBe('romanUpper');
    expect(one(t.el, 'para-bullet-start').props.value).toBe(4);
    one(t.el, 'para-bullet-scheme').props.onChange({ target: { value: 'alphaLower' } });
    one(t.el, 'para-bullet-start').props.onValue(0.2);
    expect(t.onChange.mock.calls[0][0].bullet).toMatchObject({ kind: 'auto', scheme: 'alphaLower' });
    expect(t.onChange.mock.calls[1][0].bullet).toMatchObject({ startAt: 1 });
    const plain = tree({ bullet: { kind: 'auto' } });
    expect(one(plain.el, 'para-bullet-scheme').props.value).toBe('arabic');
    expect(one(plain.el, 'para-bullet-start').props.value).toBe(1);
  });

  it('line height switches between a multiple and exact px', () => {
    const t = tree({ lineHeight: 1.5 });
    expect(one(t.el, 'para-lh').props.value).toBe(1.5);
    one(t.el, 'para-lh-mode').props.onChange({ target: { value: 'exact' } });
    one(t.el, 'para-lh-mode').props.onChange({ target: { value: 'multiple' } });
    one(t.el, 'para-lh').props.onValue(0.1);
    expect(t.onChange.mock.calls.map((c) => c[0])).toEqual([{ lineHeightPx: 24 }, { lineHeightPx: null }, { lineHeight: 0.5 }]);
    const e = tree({ lineHeightPx: 30 });
    expect(one(e.el, 'para-lh-mode').props.value).toBe('exact');
    expect(one(e.el, 'para-lh').props.value).toBe(30);
    one(e.el, 'para-lh').props.onValue(-4);
    expect(e.onChange).toHaveBeenCalledWith({ lineHeightPx: 1 });
  });

  it('spacing and margins patch their attrs, and spacing replaces a percent form', () => {
    const t = tree({ spaceBefore: 4, spaceAfterPct: 50, marginLeft: 20, firstLineIndent: -6 });
    expect(one(t.el, 'para-space-before').props.value).toBe(4);
    expect(one(t.el, 'para-space-after').props.value).toBe(0);
    expect(one(t.el, 'para-margin-left').props.value).toBe(20);
    expect(one(t.el, 'para-first-indent').props.value).toBe(-6);
    one(t.el, 'para-space-before').props.onValue(-1);
    one(t.el, 'para-space-after').props.onValue(7);
    one(t.el, 'para-margin-left').props.onValue(-9);
    one(t.el, 'para-first-indent').props.onValue(-12);
    expect(t.onChange.mock.calls.map((c) => c[0])).toEqual([
      { spaceBefore: 0, spaceBeforePct: null },
      { spaceAfter: 7, spaceAfterPct: null },
      { marginLeft: 0 },
      { firstLineIndent: -12 },
    ]);
    expect(one(tree({ spaceBeforePct: 10 }).el, 'para-space-before').props.value).toBe(0);
    const labels = JSON.stringify(tree({ spaceBeforePct: 10, spaceAfter: 5 }).el);
    expect(labels).toContain('Space before (10% of text, px replaces)');
    expect(labels).toContain('Space after (px)');
    expect(one(tree({ spaceAfter: 5, spaceBefore: 0 }).el, 'para-space-after').props.value).toBe(5);
  });
});
