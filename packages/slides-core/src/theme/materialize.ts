// Materialize a layout's placeholders into real elements. addSlide({ layoutId })
// calls this so a new slide arrives with editable title/body/etc. elements
// carrying placeholder metadata (kept for PPTX fidelity) and the layout's
// default text style baked into each element's bodyStyle.

import { createTextElement } from '../scene/elements';
import { emptyRichText } from '../text/model';
import type { TextElement } from '../scene/types';
import type { PlaceholderType } from '../scene/types';
import type { LayoutTextStyles, LevelTextStyle, PlaceholderDef } from './types';

/**
 * The effective style of one paragraph level for a placeholder type: the level's own fields over the nearest lower
 * levels' (PPTX list styles inherit level to level). Out-of-range levels clamp to 0..8.
 */
export function resolveLevelStyle(styles: LayoutTextStyles | undefined, type: PlaceholderType, level: number): LevelTextStyle {
  const levels = styles?.[type];
  if (!levels) return {};
  const top = Math.min(8, Math.max(0, Math.trunc(level) || 0), levels.length - 1);
  let out: LevelTextStyle = {};
  for (let i = 0; i <= top; i++) out = { ...out, ...levels[i] };
  return out;
}

/** Build the element for one placeholder at the given z-index. */
export function buildPlaceholderElement(
  ph: PlaceholderDef,
  containerId: string,
  index: string,
  textStyles?: LayoutTextStyles,
): TextElement {
  // The placeholder's own style wins over the layout's level-0 text defaults.
  const level0 = resolveLevelStyle(textStyles, ph.type, 0);
  const style = { ...level0, ...ph.style };
  const doc = emptyRichText();
  const paraAttrs: Record<string, unknown> = {};
  if (style.align) paraAttrs.align = style.align;
  if (level0.bullet) paraAttrs.bullet = level0.bullet;
  if (level0.marginLeft !== undefined) paraAttrs.marginLeft = level0.marginLeft;
  if (level0.firstLineIndent !== undefined) paraAttrs.firstLineIndent = level0.firstLineIndent;
  if (level0.lineHeight !== undefined) paraAttrs.lineHeight = level0.lineHeight;
  if (level0.spaceBefore !== undefined) paraAttrs.spaceBefore = level0.spaceBefore;
  if (Object.keys(paraAttrs).length) doc.content[0].attrs = { ...doc.content[0].attrs, ...paraAttrs };

  return createTextElement({
    containerId,
    index,
    x: ph.frame.x,
    y: ph.frame.y,
    w: ph.frame.w,
    h: ph.frame.h,
    rotation: ph.frame.rotation,
    placeholder: { type: ph.type, idx: ph.idx },
    richText: doc,
    bodyStyle: {
      vAlign: style.vAlign ?? 'top',
      padding: 8,
      fontSize: style.fontSize,
      color: style.color,
      bold: style.bold,
      fontFamily: style.font,
    },
  });
}
