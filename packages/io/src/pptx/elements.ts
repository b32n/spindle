import type { PptxElement } from '@b32nio/spindle-shared/export/pptx';
import type { XmlElement } from '../xlsx/xml';
import { readPicture, type PictureContext } from './pictures';
import { readShape, treeChildren, type ShapeContext } from './shapes';
import { readTable, tableOf, type TableContext } from './tables';

// The elements of a slide's shape tree in drawing order (plan K9c and K10): `p:sp` shapes, `p:pic` pictures and the table
// `p:graphicFrame`s. A `p:graphicFrame` of any other kind (chart, SmartArt, OLE: K11), `p:grpSp` and `p:cxnSp` (K10b) are not read yet.

export type ElementContext = ShapeContext & PictureContext & TableContext;

export async function readElements(tree: XmlElement | undefined, ctx: ElementContext): Promise<PptxElement[]> {
  const out: PptxElement[] = [];
  for (const c of treeChildren(tree)) {
    if (c.name === 'sp') out.push(readShape(c, ctx));
    else if (c.name === 'pic') out.push(await readPicture(c, ctx));
    else if (c.name === 'graphicFrame') {
      const tbl = tableOf(c);
      if (tbl) out.push(readTable(c, tbl, ctx));
    }
  }
  return out;
}
