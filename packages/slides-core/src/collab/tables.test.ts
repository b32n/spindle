import * as Y from 'yjs';
import { DeckImpl } from '../deck';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';
import type { Stroke, TableElement } from '../scene/types';

const RED: Stroke = { color: { kind: 'rgb', hex: '#ff0000' }, width: 2 };

function deckWithTable(): { d: DeckImpl; id: string; slide: string } {
  const d = new DeckImpl('room', 'Deck');
  const slide = d.getActiveSlideId();
  const id = d.addElement(slide, { type: 'table', rows: 3, cols: 3, x: 0, y: 0, w: 300, h: 150 }).id;
  return { d, id, slide };
}
const tableOf = (d: DeckImpl, id: string) => d.getElement(id) as TableElement;

describe('table merge and borders in the deck', () => {
  it('merge, border and split are single undoable edits', () => {
    const { d, id } = deckWithTable();
    d.mergeTableCells(id, { r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(tableOf(d, id).cells[0][0]).toMatchObject({ gridSpan: 2, rowSpan: 2 });
    d.setTableCellBorders(id, { r0: 0, c0: 0, r1: 0, c1: 0 }, 'outer', RED);
    expect(tableOf(d, id).cells[0][2].borders).toEqual({ l: RED });
    d.splitTableCell(id, 1, 1);
    expect(tableOf(d, id).cells[0][0].gridSpan).toBeUndefined();
    d.undo();
    expect(tableOf(d, id).cells[0][0].gridSpan).toBe(2);
    d.undo();
    d.undo();
    expect(tableOf(d, id).cells[0][0].gridSpan).toBeUndefined();
  });
  it('no-ops and non-tables do not touch the history', () => {
    const { d, id, slide } = deckWithTable();
    const text = d.addElement(slide, { type: 'text', x: 0, y: 0, w: 10, h: 10 }).id;
    d.mergeTableCells(id, { r0: 0, c0: 0, r1: 0, c1: 0 });
    d.splitTableCell(id, 0, 0);
    d.mergeTableCells(text, { r0: 0, c0: 0, r1: 1, c1: 1 });
    d.undo(); // undoes the text box add (the last real edit), not a merge
    expect(d.getElement(text)).toBeUndefined();
    expect(d.getElement(id)).toBeDefined();
  });
  it('removing rows and columns by range keeps spans consistent', () => {
    const { d, id } = deckWithTable();
    d.mergeTableCells(id, { r0: 0, c0: 0, r1: 1, c1: 1 });
    d.removeTableRows(id, 0, 0);
    d.removeTableColumns(id, 0, 0);
    expect(tableOf(d, id).cells.flat().some((c) => c.gridSpan || c.rowSpan || c.hMerge || c.vMerge)).toBe(false);
  });
});

describe('tables in collab (schema v9)', () => {
  it('the schema is at least 9 and a v8 client refuses a v9 room', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(9);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, new DeckImpl('d').getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
    expect(() => assertSchemaVersion(ydoc, 8)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 8)).toThrow(/up to version 8/);
    expect(() => assertSchemaVersion(ydoc, COLLAB_SCHEMA_VERSION)).not.toThrow();
  });

  it('merged cells, borders and style flags survive hydrate and serialize', () => {
    const { d, id, slide } = deckWithTable();
    d.mergeTableCells(id, { r0: 0, c0: 0, r1: 0, c1: 1 });
    d.setTableCellBorders(id, { r0: 1, c0: 1, r1: 2, c1: 2 }, 'outer', null);
    d.updateElement(id, { styleFlags: { firstRow: true, bandRow: true } });
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, slide);
    const el = back.slides[0].elements.find((e) => e.id === id) as TableElement;
    expect(el.cells).toEqual(tableOf(d, id).cells);
    expect(el.cells[0][0].gridSpan).toBe(2);
    expect(el.cells[0][1].hMerge).toBe(true);
    expect(el.cells[1][1].borders).toEqual({ l: null, t: null });
    expect(el.styleFlags).toEqual({ firstRow: true, bandRow: true });
  });

  it('a merged table is not a collab blocker', () => {
    const { d, id } = deckWithTable();
    d.mergeTableCells(id, { r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(getCollabBlockers(d.getData())).toEqual([]);
  });
});
