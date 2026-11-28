import type { PptxColorMap, PptxShadow } from '@b32nio/spindle-shared/export/pptx';
import { isOn, kid, type XmlElement } from '../xlsx/xml';
import { colorElement, parseColor } from './color';
import { EMU_PER_PX, num } from './units';

// The outer shadow of a shape (plan K6 and K13): an `a:effectLst` that holds exactly one `a:outerShdw`.

/** Scale and skew of the shadow; the model keeps none of them, so a shadow that states one stays reported as lost. */
const TRANSFORM_ATTRS = ['sx', 'sy', 'kx', 'ky'];

/** `a:effectLst` of a shape's `p:spPr` as a shadow; `undefined` when it holds anything else (glow, reflection, soft edges, an inner or preset shadow, a second shadow) or the shadow has no readable colour, so the whole list stays reported as lost. */
export function parseShadow(effectLst: XmlElement | undefined, map: PptxColorMap): PptxShadow | undefined {
  if (!effectLst || effectLst.children.length !== 1) return undefined;
  const shdw = kid(effectLst, 'outerShdw');
  const color = parseColor(colorElement(shdw), map);
  if (!shdw || !color || TRANSFORM_ATTRS.some((k) => shdw.attrs[k] !== undefined)) return undefined;
  const blur = num(shdw.attrs.blurRad);
  const dist = num(shdw.attrs.dist);
  const dir = num(shdw.attrs.dir);
  const algn = shdw.attrs.algn;
  return {
    color,
    ...(blur !== undefined && blur >= 0 ? { blur: blur / EMU_PER_PX } : {}),
    ...(dist !== undefined && dist >= 0 ? { dist: dist / EMU_PER_PX } : {}),
    ...(dir !== undefined ? { dir: dir / 60000 } : {}),
    ...(algn ? { algn } : {}),
    // ECMA-376: the attribute defaults to true, so only an absent one is left out
    ...(shdw.attrs.rotWithShape !== undefined ? { rotWithShape: isOn(shdw.attrs.rotWithShape) } : {}),
  };
}
