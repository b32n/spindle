import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { DeckImpl, type Stroke, type TableElement } from '@b32nio/spindle-slides-core';
import { DeckProvider, useDeckContext } from '../context/DeckContext';
import { SlideView } from './SlideView';
import { TableCellMenu } from './TableCellMenu';
import type { MenuItem } from './Menu';

let lastItems: Array<MenuItem | 'sep'> = [];
jest.mock('./Menu', () => ({
  Menu: ({ items }: { items: Array<MenuItem | 'sep'> }) => { lastItems = items; return null; },
}));

const RED: Stroke = { color: { kind: 'rgb', hex: '#ff0000' }, width: 3, dash: 'dash' };

function Seed({ tableId, anchor, focus }: { tableId: string; anchor: [number, number]; focus: [number, number] }): null {
  useDeckContext().tableSel.set(tableId, anchor, focus);
  return null;
}

function setup(rows = 3, cols = 3): { deck: DeckImpl; id: string; slide: string } {
  const deck = new DeckImpl('d', 'Deck');
  const slide = deck.getActiveSlideId();
  const id = deck.addElement(slide, { type: 'table', rows, cols, x: 0, y: 0, w: 300, h: 150 }).id;
  return { deck, id, slide };
}

function renderSlide(deck: DeckImpl, slide: string, seed?: { id: string; anchor: [number, number]; focus: [number, number] }): string {
  return renderToStaticMarkup(
    <DeckProvider deck={deck}>
      {seed && <Seed tableId={seed.id} anchor={seed.anchor} focus={seed.focus} />}
      <SlideView slideId={slide} interactive={!!seed} />
    </DeckProvider>,
  );
}

function menu(deck: DeckImpl, id: string, row: number, col: number, seed?: { anchor: [number, number]; focus: [number, number] }): MenuItem[] {
  renderToStaticMarkup(
    <DeckProvider deck={deck}>
      {seed && <Seed tableId={id} anchor={seed.anchor} focus={seed.focus} />}
      <TableCellMenu tableId={id} row={row} col={col} x={0} y={0} onClose={() => {}} />
    </DeckProvider>,
  );
  return lastItems.filter((i): i is MenuItem => i !== 'sep');
}
const item = (items: MenuItem[], label: string): MenuItem => items.find((i) => i.label === label)!;

describe('merged table rendering', () => {
  beforeAll(() => {
    jest.spyOn(React, 'useSyncExternalStore').mockImplementation((_sub, getSnapshot) => getSnapshot());
  });
  afterAll(() => jest.restoreAllMocks());

  it('draws the origin with colSpan/rowSpan and omits covered cells', () => {
    const { deck, id, slide } = setup();
    deck.mergeTableCells(id, { r0: 0, c0: 0, r1: 1, c1: 1 });
    const html = renderSlide(deck, slide);
    expect(html).toMatch(/<td[^>]*data-cell="0,0"[^>]*colSpan="2"[^>]*rowSpan="2"|<td[^>]*colSpan="2"[^>]*rowSpan="2"[^>]*data-cell="0,0"/);
    for (const covered of ['0,1', '1,0', '1,1']) expect(html).not.toContain(`data-cell="${covered}"`);
    expect(html).toContain('data-cell="0,2"');
    expect((html.match(/<td/g) ?? []).length).toBe(9 - 3);
  });

  it('a covered cell whose origin is missing is still drawn (no vanishing content)', () => {
    const { deck, id, slide } = setup(1, 2);
    const el = deck.getElement(id) as TableElement;
    const cells = el.cells.map((r) => r.slice());
    cells[0][1] = { ...cells[0][1], hMerge: true };
    deck.updateElement(id, { cells });
    expect(renderSlide(deck, slide)).toContain('data-cell="0,1"');
  });

  it('paints per-cell borders, dashes, no-line sides and the grid line elsewhere', () => {
    const { deck, id, slide } = setup(2, 2);
    deck.setTableCellBorders(id, { r0: 0, c0: 0, r1: 0, c1: 0 }, 'r', RED);
    deck.setTableCellBorders(id, { r0: 1, c0: 1, r1: 1, c1: 1 }, 'b', null);
    const html = renderSlide(deck, slide);
    expect(html).toContain('border-right:3px dashed rgb(255, 0, 0)');
    expect(html).toContain('border-left:3px dashed rgb(255, 0, 0)');
    expect(html).toContain('border-bottom:none');
    expect(html).toContain('border-top:1px solid');
    const dotted = setup(1, 1);
    dotted.deck.setTableCellBorders(dotted.id, { r0: 0, c0: 0, r1: 0, c1: 0 }, 't', { ...RED, dash: 'dot', width: 0 });
    expect(renderSlide(dotted.deck, dotted.slide)).toContain('border-top:1px dotted');
  });

  it('a table without a grid line falls back to the theme line', () => {
    const { deck, id, slide } = setup(1, 1);
    deck.updateElement(id, { border: undefined });
    expect(renderSlide(deck, slide)).toContain('border-left:1px solid');
  });

  it('honours per-side cell margins over uniform padding', () => {
    const { deck, id, slide } = setup(1, 2);
    deck.updateTableCells(id, [[0, 0]], { bodyStyle: { insets: { l: 1, t: 2, r: 3, b: 4 }, padding: 9 } });
    deck.updateTableCells(id, [[0, 1]], { bodyStyle: { padding: 9 } });
    const html = renderSlide(deck, slide);
    expect(html).toContain('padding:2px 3px 4px 1px');
    expect(html).toContain('padding:9px');
  });

  it('falls back to the default padding when none is set', () => {
    const { deck, slide } = setup(1, 1);
    expect(renderSlide(deck, slide)).toContain('padding:6px');
  });

  it('tints a merged cell when any covered cell is selected, and only then', () => {
    const { deck, id, slide } = setup();
    deck.mergeTableCells(id, { r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(renderSlide(deck, slide, { id, anchor: [1, 1], focus: [1, 1] })).toContain('inset 0 0 0 9999px');
    expect(renderSlide(deck, slide, { id, anchor: [2, 2], focus: [2, 2] })).toContain('inset 0 0 0 9999px'); // the unmerged (2,2)
    const merged = renderSlide(deck, slide, { id, anchor: [2, 0], focus: [2, 2] });
    expect(merged.match(/inset 0 0 0 9999px/g)).toHaveLength(3);
  });
});

describe('table cell menu: merge, split and borders', () => {
  beforeAll(() => {
    jest.spyOn(React, 'useSyncExternalStore').mockImplementation((_sub, getSnapshot) => getSnapshot());
  });
  afterAll(() => jest.restoreAllMocks());

  it('merge needs a multi-cell range; split needs a merged cell', () => {
    const { deck, id } = setup();
    let items = menu(deck, id, 0, 0);
    expect(item(items, 'Merge cells').disabled).toBe(true);
    expect(item(items, 'Split cell').disabled).toBe(true);
    items = menu(deck, id, 0, 0, { anchor: [0, 0], focus: [1, 1] });
    expect(item(items, 'Merge cells').disabled).toBe(false);
    item(items, 'Merge cells').run();
    expect((deck.getElement(id) as TableElement).cells[0][0]).toMatchObject({ gridSpan: 2, rowSpan: 2 });
    items = menu(deck, id, 1, 1);
    expect(item(items, 'Merge cells').disabled).toBe(true); // the whole region is already one cell
    expect(item(items, 'Split cell').disabled).toBe(false);
    item(items, 'Split cell').run();
    expect((deck.getElement(id) as TableElement).cells[0][0].gridSpan).toBeUndefined();
  });

  it('border commands act on the selection with the grid line, and reset', () => {
    const { deck, id } = setup();
    const items = menu(deck, id, 1, 1, { anchor: [1, 1], focus: [1, 1] });
    item(items, 'Outside borders').run();
    const t = () => deck.getElement(id) as TableElement;
    expect(t().cells[1][1].borders).toMatchObject({ l: { width: 1 }, r: { width: 1 }, t: { width: 1 }, b: { width: 1 } });
    item(items, 'No borders').run();
    expect(t().cells[1][1].borders).toEqual({ l: null, r: null, t: null, b: null });
    item(items, 'All borders').run();
    expect(t().cells[1][1].borders?.l).toMatchObject({ width: 1 });
    item(items, 'Reset borders').run();
    expect(t().cells[1][1].borders).toBeUndefined();
    deck.updateElement(id, { border: undefined });
    item(menu(deck, id, 0, 0), 'All borders').run();
    expect(t().cells[0][0].borders?.t).toMatchObject({ color: { kind: 'theme', slot: 'dk2' } });
  });
});
