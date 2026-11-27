import type { PptxConnection, PptxConnector } from '@b32nio/spindle-shared/export/pptx';
import { isTrue, kid, type XmlElement } from '../xlsx/xml';
import { parseGeometry } from './fill';
import { frameOf } from './levels';
import { shapeOutline, type ShapeContext } from './shapes';
import { num } from './units';

// Connectors (plan K10b): `p:cxnSp` with its route (the preset: straight, bent or curved, with its adjustments), outline and arrow
// heads, and the shapes its ends are glued to (`a:stCxn`, `a:endCxn`).

/** A connection end: the id of the shape and a non-negative site index; anything else is no connection. */
function readConnection(e: XmlElement | undefined): PptxConnection | undefined {
  const idx = num(e?.attrs.idx);
  if (!e?.attrs.id || idx === undefined || idx < 0 || !Number.isInteger(idx)) return undefined;
  return { id: e.attrs.id, idx };
}

export function readConnector(e: XmlElement, ctx: ShapeContext): PptxConnector {
  const nv = kid(e, 'nvCxnSpPr');
  const cNvPr = kid(nv, 'cNvPr');
  const cNv = kid(nv, 'cNvCxnSpPr');
  const spPr = kid(e, 'spPr');
  const frame = frameOf(kid(spPr, 'xfrm'));
  const geometry = parseGeometry(spPr);
  const line = shapeOutline(spPr, kid(e, 'style'), ctx);
  const start = readConnection(kid(cNv, 'stCxn'));
  const end = readConnection(kid(cNv, 'endCxn'));
  const altText = (cNvPr?.attrs.descr ?? '').trim();
  return {
    kind: 'connector',
    pptxId: cNvPr?.attrs.id ?? '',
    name: (cNvPr?.attrs.name ?? '').trim(),
    ...(altText !== '' ? { altText } : {}),
    ...(isTrue(cNvPr?.attrs.hidden) ? { hidden: true } : {}),
    ...(frame ? { frame } : {}),
    ...(geometry ? { geometry } : {}),
    ...(line ? { line } : {}),
    ...(start ? { start } : {}),
    ...(end ? { end } : {}),
  };
}
