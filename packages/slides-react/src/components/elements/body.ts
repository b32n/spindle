// Body-level text layout shared by the static renderer, the live editor and shapes: insets, anchor, vertical text,
// columns, and the autofit scale variables. One implementation so edit mode lays text out exactly as idle mode does.

import type React from 'react';
import {
  autofitScales,
  hasAutofit,
  insetsToCss,
  resolveColumns,
  resolveInsets,
  textVert,
  writingModeCss,
  type BodyStyle,
} from '@b32nio/spindle-slides-core';

const V_ALIGN = { top: 'flex-start', middle: 'center', bottom: 'flex-end' } as const;

/** The CSS variables the autofit scales travel in (the editor schema's font sizes and line heights read them). */
export function scaleVars(body: BodyStyle | undefined): Record<string, string> {
  const s = autofitScales(body);
  return { '--spindle-fs': String(s.font), '--spindle-ls': String(s.line) };
}

/** Column variables for the live editor's stylesheet (`.ProseMirror` reads them). */
export function columnVars(body: BodyStyle | undefined): Record<string, string> {
  const c = resolveColumns(body);
  return { '--spindle-cols': String(c.count), '--spindle-colgap': `${c.gap}px` };
}

/**
 * The box style of a text body: padding from the insets, vertical anchor, centred text block, writing mode.
 * `defaultVAlign` is the anchor when the body names none (shapes centre their text).
 */
export function bodyBoxStyle(body: BodyStyle | undefined, defaultVAlign: keyof typeof V_ALIGN = 'top'): React.CSSProperties {
  return {
    boxSizing: 'border-box',
    padding: insetsToCss(resolveInsets(body)),
    display: 'flex',
    flexDirection: 'column',
    justifyContent: V_ALIGN[body?.vAlign ?? defaultVAlign] ?? V_ALIGN[defaultVAlign],
    ...(body?.anchorCtr ? { alignItems: 'center' } : {}),
    ...writingModeCss(textVert(body)),
  };
}

/** Whether the live editor should let its text block shrink to the content so `anchorCtr` can centre it (not with columns). */
export function centresBlock(body: BodyStyle | undefined): boolean {
  return !!body?.anchorCtr && resolveColumns(body).count < 2;
}

/** Overflow of the element box: a body that declares autofit shows its text in full; an unset body keeps clipping. */
export function bodyOverflow(body: BodyStyle | undefined): 'hidden' | 'visible' {
  return hasAutofit(body) ? 'visible' : 'hidden';
}

/** Style for the column wrapper (only rendered when the body has two or more columns). */
export function columnsStyle(body: BodyStyle | undefined): React.CSSProperties | undefined {
  const c = resolveColumns(body);
  return c.count > 1 ? { columnCount: c.count, columnGap: c.gap, width: '100%', flexShrink: 0 } : undefined;
}
