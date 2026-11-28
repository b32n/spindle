// Layout and theme mappings between the PPTX interchange and the slides model. The reader flattens master -> layout inheritance
// (plan decision SD6), so a layout here is the flat `PptxLayout`: its placeholders with effective frames and text levels, and its
// static shapes. The model keeps text levels per placeholder *type* (`LayoutData.textStyles`), so two placeholders of one type on a
// layout share the first one's levels.

import type {
  PptxColorMap,
  PptxLayout,
  PptxLevelStyle,
  PptxPlaceholder,
  PptxShape,
  PptxTheme,
} from '@b32nio/spindle-shared/export/pptx';
import { indexesBetween } from '../scene/fractional-index';
import type { PlaceholderType, SlideElement } from '../scene/types';
import type { LayoutData, LayoutTextStyles, LevelTextStyle, PlaceholderDef, ThemeData } from '../theme/types';
import { resolveLevelStyle } from '../theme/materialize';
import { elementsFromPptx, elementsToPptx, placeholderTypeFromPptx, placeholderTypeToPptx, type ElementTally } from './pptx-elements';
import { bulletFromPptx } from './pptx-text';
import { colorFromPptx, colorToPptx, finite, ptToPx, pxToPt } from './pptx-values';
import type { Bullet } from '../text/model';

export const MASTER_PART = 'ppt/slideMasters/slideMaster1.xml';

export const DEFAULT_COLOR_MAP: PptxColorMap = {
  bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2',
  accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
  hlink: 'hlink', folHlink: 'folHlink',
};

export const themeFromPptx = (t: PptxTheme): ThemeData => ({ name: t.name, colors: { ...t.colors }, fonts: { major: t.fonts.major, minor: t.fonts.minor } });
export const themeToPptx = (t: ThemeData): PptxTheme => ({ name: t.name, colors: { ...t.colors }, fonts: { major: t.fonts.major, minor: t.fonts.minor } });

function levelFromPptx(s: PptxLevelStyle): LevelTextStyle {
  const out: LevelTextStyle = {};
  if (s.color) out.color = colorFromPptx(s.color);
  if (s.fontSize !== undefined) out.fontSize = ptToPx(s.fontSize);
  if (s.bold !== undefined) out.bold = s.bold;
  if (s.italic !== undefined) out.italic = s.italic;
  if (s.align) out.align = s.align;
  if (s.font !== undefined) out.font = s.font;
  if (s.bullet) out.bullet = bulletFromPptx(s.bullet, s);
  if (s.marginLeft !== undefined) out.marginLeft = s.marginLeft;
  if (s.indent !== undefined) out.firstLineIndent = s.indent;
  if (s.lineHeight !== undefined) out.lineHeight = s.lineHeight;
  if (s.spaceBefore !== undefined) out.spaceBefore = ptToPx(s.spaceBefore);
  return out;
}

const BULLET_AUTO: Record<NonNullable<Bullet['scheme']>, string> = { arabic: 'arabicPeriod', alphaLower: 'alphaLcPeriod', alphaUpper: 'alphaUcPeriod', romanLower: 'romanLcPeriod', romanUpper: 'romanUcPeriod' };

function levelToPptx(s: LevelTextStyle): PptxLevelStyle {
  const out: PptxLevelStyle = {};
  if (s.color) out.color = colorToPptx(s.color);
  const size = finite(s.fontSize);
  if (size !== undefined && size > 0) out.fontSize = pxToPt(size);
  if (s.bold !== undefined) out.bold = s.bold;
  if (s.italic !== undefined) out.italic = s.italic;
  if (s.align) out.align = s.align;
  if (s.font !== undefined) out.font = s.font;
  const b = s.bullet;
  if (b) {
    out.bullet = b.kind === 'none' ? { kind: 'none' } : b.kind === 'char' ? { kind: 'char', char: b.char ?? '•' } : { kind: 'auto', scheme: BULLET_AUTO[b.scheme ?? 'arabic'], ...(finite(b.startAt) !== undefined ? { startAt: b.startAt } : {}) };
    if (b.kind !== 'none') {
      if (b.color) out.bulletColor = colorToPptx(b.color);
      const pct = finite(b.sizePct);
      if (pct !== undefined) out.bulletSizePct = pct / 100;
    }
  }
  if (finite(s.marginLeft) !== undefined) out.marginLeft = s.marginLeft;
  if (finite(s.firstLineIndent) !== undefined) out.indent = s.firstLineIndent;
  if (finite(s.lineHeight) !== undefined) out.lineHeight = s.lineHeight;
  if (finite(s.spaceBefore)) out.spaceBefore = pxToPt(s.spaceBefore as number);
  return out;
}

function placeholderFromPptx(p: PptxPlaceholder): { def: PlaceholderDef; type: PlaceholderType } | undefined {
  const type = placeholderTypeFromPptx(p.type);
  if (!type || !p.frame) return undefined;
  const vAlign = p.anchor ?? p.bodyProps?.anchor;
  return {
    type,
    def: {
      type,
      idx: p.idx ?? 0,
      frame: {
        x: p.frame.x,
        y: p.frame.y,
        w: p.frame.w,
        h: p.frame.h,
        rotation: p.frame.rot ?? 0,
      },
      ...(p.prompt ? { prompt: p.prompt } : {}),
      ...(vAlign ? { style: { vAlign } } : {}),
    },
  };
}

/** A layout as the model keeps it. `id` is the layout's part name, which slides name in `layoutRef` and `PptxSlide.layoutPart`. */
export function layoutFromPptx(l: PptxLayout, tally: ElementTally): LayoutData {
  const placeholders: PlaceholderDef[] = [];
  const textStyles: LayoutTextStyles = {};
  for (const p of l.placeholders) {
    const mapped = placeholderFromPptx(p);
    if (!mapped) continue;
    placeholders.push(mapped.def);
    if (!textStyles[mapped.type]) {
      const levels = p.levels.map(levelFromPptx);
      while (levels.length > 0 && Object.keys(levels[levels.length - 1]).length === 0) levels.pop();
      if (levels.length > 0) textStyles[mapped.type] = levels;
    }
  }
  const elements = elementsFromPptx(l.elements, l.partName, tally);
  return {
    id: l.partName,
    name: l.name,
    ...(l.masterName !== undefined ? { masterName: l.masterName } : {}),
    ...(l.background ? { background: { kind: 'solid', color: colorFromPptx(l.background) } } : {}),
    placeholders,
    ...(elements.length > 0 ? { elements } : {}),
    ...(Object.keys(textStyles).length > 0 ? { textStyles } : {}),
  };
}

export function layoutToPptx(l: LayoutData): PptxLayout {
  const placeholders: PptxPlaceholder[] = l.placeholders.map((p) => {
    // the placeholder's own style wins over the layout's first-level text defaults, as when a slide is made from it
    const first = { ...resolveLevelStyle(l.textStyles, p.type, 0), ...p.style };
    const levels: PptxLevelStyle[] = Array.from({ length: 9 }, (_, i) => levelToPptx(i === 0 ? first : resolveLevelStyle(l.textStyles, p.type, i)));
    const vAlign = p.style?.vAlign;
    return {
      type: placeholderTypeToPptx(p.type),
      idx: p.idx,
      frame: { x: p.frame.x, y: p.frame.y, w: p.frame.w, h: p.frame.h, ...(p.frame.rotation ? { rot: p.frame.rotation } : {}) },
      ...(vAlign ? { anchor: vAlign } : {}),
      levels,
      ...(p.prompt ? { prompt: p.prompt } : {}),
    };
  });
  const elements = elementsToPptx(withIndex(l.elements ?? [])).elements.filter((e): e is PptxShape => e.kind === 'shape');
  const bg = l.background?.kind === 'solid' ? colorToPptx(l.background.color) : undefined;
  return {
    partName: l.id,
    name: l.name,
    masterPart: MASTER_PART,
    ...(l.masterName !== undefined ? { masterName: l.masterName } : {}),
    ...(bg ? { background: bg } : {}),
    placeholders,
    elements,
    staticShapes: elements.length,
  };
}

/** Layout elements the editor's own tools never made have no z-order key. */
function withIndex(elements: SlideElement[]): SlideElement[] {
  const keys = indexesBetween(null, null, elements.length);
  return elements.map((e, i) => ({ ...e, index: e.index || keys[i] }));
}
