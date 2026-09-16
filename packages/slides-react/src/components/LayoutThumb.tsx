// LayoutThumb — a small SVG preview of a layout, drawn from its placeholder
// frames (like the layout gallery in PowerPoint / Google Slides). Title-ish
// placeholders render as a solid bar; body placeholders as a few text lines.

import React from 'react';
import type { LayoutData, PlaceholderDef } from '@b32nio/spindle-slides-core';

const BAR = '#aeb9cc';
const SUBTLE = '#cdd5e1';
const LINE = '#d7dde6';

function Placeholder({ p }: { p: PlaceholderDef }): React.ReactElement {
  const { x, y, w, h } = p.frame;
  if (p.type === 'body') {
    // A few evenly-spaced text lines filling the frame.
    const rows = Math.max(2, Math.min(6, Math.round(h / 90)));
    const gap = h / rows;
    return (
      <g>
        {Array.from({ length: rows }, (_, i) => (
          <rect key={i} x={x} y={y + gap * i + gap * 0.25} width={i === rows - 1 ? w * 0.6 : w} height={Math.min(18, gap * 0.4)} rx={6} fill={LINE} />
        ))}
      </g>
    );
  }
  const fill = p.type === 'subtitle' ? SUBTLE : BAR;
  return <rect x={x} y={y} width={w} height={h} rx={10} fill={fill} />;
}

export function LayoutThumb({ layout, size }: { layout: LayoutData; size: { w: number; h: number } }): React.ReactElement {
  return (
    <svg viewBox={`0 0 ${size.w} ${size.h}`} width="100%" height="100%" style={{ display: 'block' }} preserveAspectRatio="xMidYMid meet">
      <rect x={0} y={0} width={size.w} height={size.h} fill="#fff" />
      {layout.placeholders.map((p) => (
        <Placeholder key={`${p.type}-${p.idx}`} p={p} />
      ))}
    </svg>
  );
}
