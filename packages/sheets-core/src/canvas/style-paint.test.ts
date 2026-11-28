import { BORDER_PAINT, INDENT_PX, borderKindOf, fillColorOf } from './style-paint';
import { GridRenderer } from './grid-renderer';
import { CellRenderer } from './cell-renderer';
import { TextRenderer } from './text-renderer';
import { CSS_OF_KIND } from '../border-kinds';
import type { BorderKind } from '../types';
import { DEFAULT_THEME } from './types';

const KINDS = Object.keys(CSS_OF_KIND) as BorderKind[];

/** A canvas context that records what is drawn. */
function recorder() {
  const calls: Array<[string, ...unknown[]]> = [];
  const state: Record<string, unknown> = { font: '', textAlign: '' };
  const ctx = new Proxy(state, {
    get: (t, k: string) => (k in t ? t[k] : (...args: unknown[]) => { calls.push([k, ...args]); return k === 'measureText' ? { width: 10 * String(args[0]).length, actualBoundingBoxAscent: 8 } : undefined; }),
    set: (t, k: string, v) => { t[k] = v; if (k !== 'font' && k !== 'textAlign') calls.push([`set ${k}`, v]); return true; },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, calls, named: (n: string) => calls.filter((c) => c[0] === n) };
}

beforeAll(() => {
  // the text renderer measures on an offscreen canvas: a stub with a measuring context
  (globalThis as unknown as { document: unknown }).document = { createElement: () => ({ getContext: () => recorder().ctx }) };
});

describe('border strokes', () => {
  it('gives each of the thirteen kinds its own stroke, matching the width the editor stores', () => {
    const seen = new Set(KINDS.map((k) => JSON.stringify(BORDER_PAINT[k])));
    expect(seen.size).toBe(13);
    for (const k of KINDS) expect(BORDER_PAINT[k].width).toBe(CSS_OF_KIND[k].width);
    expect(KINDS.filter((k) => BORDER_PAINT[k].double)).toEqual(['double']);
  });

  it('takes the kind from the shorthand and the stored kind while they agree', () => {
    expect(borderKindOf('1px solid #000', undefined)).toBe('thin');
    expect(borderKindOf('1px dashed #000', 'dashDot')).toBe('dashDot');
    expect(borderKindOf('3px solid #000', 'dashDot')).toBe('thick');
    for (const bad of [undefined, '', 'none', 'solid', 'x solid #000']) expect(borderKindOf(bad, 'thin')).toBeUndefined();
  });

  it.each(KINDS)('strokes %s with its dash pattern and colour', (kind) => {
    const { ctx, named, calls } = recorder();
    const grid = new GridRenderer(DEFAULT_THEME);
    const css = CSS_OF_KIND[kind];
    const side = grid.parseBorderString(`${css.width}px ${css.style} #123456`, kind);
    expect(side?.kind).toBe(kind);
    grid.renderCellBorders(ctx, { x: 0, y: 0, width: 50, height: 20 }, { top: side });
    expect(named('stroke')).toHaveLength(BORDER_PAINT[kind].double ? 2 : 1);
    expect(calls.find((c) => c[0] === 'setLineDash')?.[1]).toEqual([...BORDER_PAINT[kind].dash]);
    expect(calls.find((c) => c[0] === 'set strokeStyle')?.[1]).toBe('#123456');
    expect(calls.find((c) => c[0] === 'set lineWidth')?.[1]).toBe(BORDER_PAINT[kind].double ? 1 : css.width);
  });

  it('draws the sides of a double border either side of the edge, horizontally and vertically', () => {
    const { ctx, named } = recorder();
    const grid = new GridRenderer(DEFAULT_THEME);
    const dbl = grid.parseBorderString('3px double #000', 'double');
    grid.renderCellBorders(ctx, { x: 0, y: 0, width: 50, height: 20 }, { top: dbl, left: dbl });
    expect(named('moveTo').map((c) => [c[1], c[2]])).toEqual([[0.5, -0.5], [0.5, 1.5], [-0.5, 0.5], [1.5, 0.5]]);
  });

  it('falls back to the CSS style when a side has no kind', () => {
    const { ctx, calls } = recorder();
    const grid = new GridRenderer(DEFAULT_THEME);
    grid.renderCellBorders(ctx, { x: 0, y: 0, width: 5, height: 5 }, { top: { width: 1, color: '#000', style: 'dashed' }, right: { width: 1, color: '#000', style: 'dotted' }, bottom: { width: 1, color: '#000', style: 'solid' }, left: { width: 2, color: '#000', style: 'double' } });
    // each side sets its dash and resets it after: dashed, dotted, solid, and the double's solid
    expect(calls.filter((c) => c[0] === 'setLineDash').map((c) => c[1])).toEqual([[4, 2], [], [1, 1], [], [], [], [], []]);
    expect(calls.filter((c) => c[0] === 'stroke')).toHaveLength(5);
  });
});

describe('pattern fills', () => {
  it('is the solid colour, or none, without a pattern', () => {
    expect(fillColorOf(undefined)).toBeUndefined();
    expect(fillColorOf({})).toBeUndefined();
    expect(fillColorOf({ backgroundColor: '#abcdef' })).toBe('#abcdef');
    expect(fillColorOf({ fillPattern: 'darkGray' })).toBeUndefined();
  });

  it('mixes the pattern colour into the colour behind it by the pattern\'s coverage', () => {
    expect(fillColorOf({ fillPattern: 'mediumGray', patternColor: '#000000' })).toBe('#808080');
    expect(fillColorOf({ fillPattern: 'darkGray', patternColor: '#000000', backgroundColor: '#ffffff' })).toBe('#404040');
    expect(fillColorOf({ fillPattern: 'gray125', patternColor: '#ff0000', backgroundColor: '#000' })).toBe('#200000');
    expect(fillColorOf({ fillPattern: 'unknownPattern', patternColor: '#000000' })).toBe('#808080');
    expect(fillColorOf({ fillPattern: 'lightGray', patternColor: '#fff' }, '#000000')).toBe('#404040');
  });

  it('uses the pattern colour when a colour cannot be mixed', () => {
    expect(fillColorOf({ fillPattern: 'darkGray', patternColor: 'red', backgroundColor: '#ffffff' })).toBe('red');
    expect(fillColorOf({ fillPattern: 'darkGray', patternColor: '#000000', backgroundColor: 'blue' })).toBe('#000000');
  });

  it('is what a cell is filled with', () => {
    const { ctx, calls } = recorder();
    new CellRenderer(DEFAULT_THEME, new TextRenderer(DEFAULT_THEME)).renderCell(ctx, undefined, { x: 0, y: 0, width: 10, height: 10 }, { fillPattern: 'mediumGray', patternColor: '#000000' }, undefined);
    expect(calls.find((c) => c[0] === 'set fillStyle')?.[1]).toBe('#808080');
  });
});

describe('text style', () => {
  const text = (style: Record<string, unknown>, bounds = { x: 0, y: 0, width: 200, height: 20 }, value = 'hello') => {
    const r = recorder();
    const tr = new TextRenderer(DEFAULT_THEME);
    tr.renderText(r.ctx, value, bounds, style);
    return r;
  };

  it('moves left-aligned text in by the indent, right-aligned text from the right, and leaves centred text', () => {
    expect(text({ indent: 2 }).named('fillText')[0][2]).toBe(4 + 2 * INDENT_PX);
    expect(text({ indent: 1, textAlign: 'right' }).named('fillText')[0][2]).toBe(200 - 4 - INDENT_PX);
    expect(text({ indent: 3, textAlign: 'center' }).named('fillText')[0][2]).toBe(100);
    expect(text({ indent: 0 }).named('fillText')[0][2]).toBe(4);
  });

  it('shrinks text that is wider than the cell, never enlarges, and leaves rotated text alone', () => {
    const fontOf = (style: Record<string, unknown>, width: number, value = 'hello') => {
      const r = text(style, { x: 0, y: 0, width, height: 20 }, value);
      return String(r.calls.find((c) => c[0] === 'fillText') && '');
    };
    void fontOf;
    const sizeIn = (style: Record<string, unknown>, width: number) => {
      const fonts: string[] = [];
      const r = recorder();
      const proxy = new Proxy(r.ctx, { set: (t, k, v) => { if (k === 'font') fonts.push(String(v)); (t as never as Record<string, unknown>)[k as string] = v; return true; } });
      new TextRenderer(DEFAULT_THEME).renderText(proxy, 'hello world', { x: 0, y: 0, width, height: 20 }, style);
      return Number(/(\d+(?:\.\d+)?)px/.exec(fonts[fonts.length - 1])?.[1]);
    };
    expect(sizeIn({ fontSize: 20 }, 60)).toBe(20);
    expect(sizeIn({ fontSize: 20, shrinkToFit: true }, 60)).toBeLessThan(20);
    expect(sizeIn({ fontSize: 20, shrinkToFit: true }, 1000)).toBe(20);
    expect(sizeIn({ fontSize: 20, shrinkToFit: true }, 8)).toBeGreaterThanOrEqual(1);
    expect(sizeIn({ fontSize: 20, shrinkToFit: true, textRotation: 30 }, 60)).toBe(20);
  });

  it('draws underline, double underline, strike, and both together', () => {
    const lines = (style: Record<string, unknown>) => text(style).named('lineTo').length;
    expect(lines({})).toBe(0);
    expect(lines({ textDecoration: 'none' })).toBe(0);
    expect(lines({ textDecoration: 'underline' })).toBe(1);
    expect(lines({ textDecoration: 'underline', underlineKind: 'double' })).toBe(2);
    expect(lines({ textDecoration: 'underline', underlineKind: 'doubleAccounting' })).toBe(2);
    expect(lines({ textDecoration: 'underline', underlineKind: 'singleAccounting' })).toBe(1);
    expect(lines({ textDecoration: 'line-through' })).toBe(1);
    expect(lines({ textDecoration: 'underline', strikethrough: true })).toBe(2);
    expect(lines({ strikethrough: true })).toBe(1);
  });

  it('decorates rotated text too', () => {
    expect(text({ textRotation: 45, textDecoration: 'underline', strikethrough: true }).named('lineTo')).toHaveLength(2);
  });
});

describe('cells', () => {
  it('draws the borders of an empty cell and passes the text style flags on', () => {
    const { ctx, named } = recorder();
    const cr = new CellRenderer(DEFAULT_THEME, new TextRenderer(DEFAULT_THEME));
    cr.renderEmptyCell(ctx, { x: 0, y: 0, width: 10, height: 10 }, { borderTop: '1px solid #000', backgroundColor: '#eee' });
    expect(named('stroke')).toHaveLength(1);
    cr.renderEmptyCell(ctx, { x: 0, y: 0, width: 10, height: 10 });
    expect(named('stroke')).toHaveLength(1);
    cr.renderCell(ctx, { value: 'abc' }, { x: 0, y: 0, width: 100, height: 20 }, { shrinkToFit: true, indent: 1, textDecoration: 'underline', underlineKind: 'double', strikethrough: true }, undefined);
    expect(named('fillText')).toHaveLength(1);
    cr.renderHyperlinkCell(ctx, { value: 'abc' }, { x: 0, y: 0, width: 100, height: 20 }, { borderBottom: '1px solid #000' }, undefined);
    expect(named('stroke').length).toBeGreaterThan(2);
  });
});
