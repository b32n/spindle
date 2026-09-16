import React from 'react';
import { resolveColor, type LineElement, type ArrowHead, type ThemeData } from '@b32nio/spindle-slides-core';
import { strokeAttrs } from './style';

function ArrowMarker({ id, head, color, place }: { id: string; head: ArrowHead; color: string; place: 'start' | 'end' }): React.ReactElement | null {
  if (head === 'none') return null;
  const path =
    head === 'triangle'
      ? 'M0 0 L10 5 L0 10 Z'
      : head === 'arrow'
        ? 'M0 0 L10 5 L0 10'
        : null;
  return (
    <marker
      id={id}
      viewBox="0 0 10 10"
      refX={place === 'end' ? 9 : 1}
      refY={5}
      markerWidth={7}
      markerHeight={7}
      orient="auto-start-reverse"
    >
      {head === 'circle' ? (
        <circle cx={5} cy={5} r={4} fill={color} />
      ) : (
        <path d={path!} fill={head === 'triangle' ? color : 'none'} stroke={color} strokeWidth={1.5} />
      )}
    </marker>
  );
}

// Vertical padding (slide units) added above and below the line's box so the
// SVG has a real, hit-testable height even for a horizontal line (box height 0).
// The element wrapper stays el.w×el.h; only this inner SVG overflows.
const PAD = 12;

/** Explicit start/end in box-local coordinates (0..w, 0..h). */
export interface LineEndpoints {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export function LineView({ el, theme, endpoints }: { el: LineElement; theme: ThemeData; endpoints?: LineEndpoints }): React.ReactElement {
  const stroke = strokeAttrs(el.stroke, theme) ?? {
    stroke: resolveColor(el.stroke.color, theme),
    strokeWidth: Math.max(1, el.stroke.width),
  };

  // Draw in absolute user units (no viewBox) so the stroke keeps a constant
  // width regardless of the box aspect — a viewBox with preserveAspectRatio
  // "none" would scale a horizontal line's stroke to nothing. The box occupies
  // y ∈ [PAD, PAD+h]; the SVG is offset up by PAD so it centres on the box.
  const w = el.w;
  const h = el.h;
  const top = PAD;
  const bottom = PAD + h;
  // A connector passes explicit endpoints so start→end direction (hence the
  // arrowhead) is exact; a free line reads its start corner from flipH/flipV,
  // which together pick one of the four orderings — so the arrow lands on the
  // right end even for a right-to-left or bottom-to-top line.
  const [x1, y1, x2, y2] = endpoints
    ? [endpoints.x1, endpoints.y1 + PAD, endpoints.x2, endpoints.y2 + PAD]
    : [el.flipH ? w : 0, el.flipV ? bottom : top, el.flipH ? 0 : w, el.flipV ? top : bottom];

  const startId = `arw-s-${el.id}`;
  const endId = `arw-e-${el.id}`;
  const startHead = el.startArrow ?? 'none';
  const endHead = el.endArrow ?? 'none';

  return (
    <svg
      width={w}
      height={h + 2 * PAD}
      style={{ position: 'absolute', left: 0, top: -PAD, overflow: 'visible', pointerEvents: 'none' }}
    >
      <defs>
        <ArrowMarker id={startId} head={startHead} color={stroke.stroke} place="start" />
        <ArrowMarker id={endId} head={endHead} color={stroke.stroke} place="end" />
      </defs>
      {/* Wide transparent hit-stroke so the thin line is easy to click. */}
      <line x1={x1} y1={y1} x2={x2} y2={y2} stroke="transparent" strokeWidth={Math.max(14, stroke.strokeWidth + 12)} strokeLinecap="round" style={{ pointerEvents: 'stroke' }} />
      <line
        x1={x1}
        y1={y1}
        x2={x2}
        y2={y2}
        stroke={stroke.stroke}
        strokeWidth={stroke.strokeWidth}
        strokeDasharray={stroke.strokeDasharray}
        strokeLinecap="round"
        markerStart={startHead !== 'none' ? `url(#${startId})` : undefined}
        markerEnd={endHead !== 'none' ? `url(#${endId})` : undefined}
      />
    </svg>
  );
}
