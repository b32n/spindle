import { Node as PmNode } from 'prosemirror-model';
import { docsSchema as schema } from '@b32nio/spindle-docs-core';
import { constrainSelectionToIsolatingBlock } from './selection-constraint';

// doc: paragraph("before") + table[ row[ cell("AAA"), cell("BBB") ] ] + paragraph("after")
const cell = (text: string): PmNode =>
  schema.nodes.table_cell.create(null, schema.nodes.paragraph.create(null, schema.text(text)));
const doc = schema.nodes.doc.create(null, [
  schema.nodes.paragraph.create(null, schema.text('before')),
  schema.nodes.table.create(null, schema.nodes.table_row.create(null, [cell('AAA'), cell('BBB')])),
  schema.nodes.paragraph.create(null, schema.text('after')),
]);

/** First position where `text` occurs inside the doc (position before the text). */
const posOf = (text: string): number => {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found === -1 && node.isText && node.text?.includes(text)) {
      found = pos + (node.text.indexOf(text));
    }
    return found === -1;
  });
  if (found === -1) throw new Error(`text ${text} not found`);
  return found;
};

describe('constrainSelectionToIsolatingBlock (H1: cross-cell/table selection corruption)', () => {
  it('clamps a head in another cell back into the anchor cell', () => {
    const anchor = posOf('AAA') + 1; // inside cell A
    const to = posOf('BBB') + 2; // inside cell B
    const constrained = constrainSelectionToIsolatingBlock(doc, anchor, to);
    // The constrained head must resolve into the SAME cell as the anchor.
    const cellOf = (p: number) => {
      const $p = doc.resolve(p);
      for (let d = $p.depth; d > 0; d--) if ($p.node(d).type.name === 'table_cell') return $p.node(d);
      return null;
    };
    expect(cellOf(constrained)).toBe(cellOf(anchor));
    expect(cellOf(constrained)).not.toBeNull();
    expect(constrained).toBeLessThan(to);
  });

  it('clamps a head outside the table back into the anchor cell', () => {
    const anchor = posOf('AAA') + 1; // inside cell A
    const to = posOf('after') + 2; // outside the table entirely
    const constrained = constrainSelectionToIsolatingBlock(doc, anchor, to);
    const $c = doc.resolve(constrained);
    // Still inside the table (never crosses the isolating boundary).
    let inTable = false;
    for (let d = $c.depth; d > 0; d--) if ($c.node(d).type.name === 'table') inTable = true;
    expect(inTable).toBe(true);
  });

  it('clamps a head entering a cell back out when the anchor is outside', () => {
    const anchor = posOf('before') + 2; // outside the table
    const to = posOf('BBB') + 1; // inside cell B
    const constrained = constrainSelectionToIsolatingBlock(doc, anchor, to);
    const $c = doc.resolve(constrained);
    let inTable = false;
    for (let d = $c.depth; d > 0; d--) if ($c.node(d).type.name === 'table') inTable = true;
    expect(inTable).toBe(false);
  });

  it('leaves an in-cell selection unchanged', () => {
    const anchor = posOf('AAA') + 1;
    const to = posOf('AAA') + 3; // same cell
    expect(constrainSelectionToIsolatingBlock(doc, anchor, to)).toBe(to);
  });

  it('leaves a plain same-block selection unchanged', () => {
    const anchor = posOf('before') + 1;
    const to = posOf('before') + 4;
    expect(constrainSelectionToIsolatingBlock(doc, anchor, to)).toBe(to);
  });
});
