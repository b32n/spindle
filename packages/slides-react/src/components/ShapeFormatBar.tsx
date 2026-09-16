// ShapeFormatBar — inline fill & border controls, rendered in the main Toolbar
// row (like TextFormatBar / LineFormatBar) when a single shape is selected. The
// model already carries `fill` and `stroke`; this is the UI for them.

import React from 'react';
import { resolveColor, type Color, type ShapeElement, type StrokeDash, type ThemeData } from '@b32nio/spindle-slides-core';
import { useDeck, useSelection, useElement, useTheme } from '../hooks';
import { ToolbarButton, ToolbarDivider } from './toolbarUI';

const field: React.CSSProperties = {
  height: 30, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 8,
  background: 'rgba(255,255,255,0.8)', color: '#334155', padding: '0 6px',
  fontSize: 13, fontFamily: '"Inter", sans-serif', cursor: 'pointer',
};
const swatch: React.CSSProperties = { ...field, width: 34, padding: 3 };
const labelWrap: React.CSSProperties = { display: 'inline-flex', alignItems: 'center', gap: 5 };
const tag: React.CSSProperties = { fontSize: 11, fontWeight: 600, letterSpacing: '.04em', color: '#64748b', textTransform: 'uppercase', fontFamily: '"Inter", sans-serif' };

const WIDTHS = [1, 2, 3, 4, 6, 8];
const DASHES: Array<{ v: StrokeDash; label: string }> = [
  { v: 'solid', label: 'Solid' },
  { v: 'dash', label: 'Dashed' },
  { v: 'dot', label: 'Dotted' },
];

/** Best-effort hex for seeding a color input (theme colors resolve to hex). */
function colorToHex(color: Color, theme: ThemeData): string {
  const s = resolveColor(color, theme);
  return /^#[0-9a-fA-F]{6}$/.test(s) ? s : '#888888';
}

export function ShapeFormatBar(): React.ReactElement | null {
  const deck = useDeck();
  const theme = useTheme();
  const selection = useSelection();
  const id = selection.elementIds.length === 1 ? selection.elementIds[0] : null;
  const el = useElement(id ?? '') as ShapeElement | undefined;
  if (!id || !el || el.type !== 'shape') return null;

  const hasFill = el.fill.kind === 'solid';
  const fillHex = el.fill.kind === 'solid' ? colorToHex(el.fill.color, theme) : '#cccccc';
  const stroke = el.stroke;
  const strokeHex = stroke ? colorToHex(stroke.color, theme) : '#334155';

  const setFill = (color: string) => deck.updateElement(id, { fill: { kind: 'solid', color: { kind: 'rgb', hex: color } } });
  const noFill = () => deck.updateElement(id, { fill: { kind: 'none' } });
  const setStroke = (patch: Partial<{ color: Color; width: number; dash: StrokeDash }>) =>
    deck.updateElement(id, { stroke: { color: { kind: 'rgb', hex: strokeHex }, width: 2, ...stroke, ...patch } });
  const noBorder = () => deck.updateElement(id, { stroke: undefined });

  return (
    <>
      <ToolbarDivider />
      <span style={labelWrap}>
        <span style={tag}>Fill</span>
        <input
          type="color"
          title="Fill color"
          value={fillHex}
          onMouseDown={(e) => e.preventDefault()}
          onChange={(e) => setFill(e.target.value)}
          style={{ ...swatch, opacity: hasFill ? 1 : 0.4 }}
        />
      </span>
      <ToolbarButton title="No fill" active={!hasFill} onClick={noFill}>
        <span style={{ fontSize: 13, fontWeight: 700 }}>⦸</span>
      </ToolbarButton>

      <ToolbarDivider />
      <span style={labelWrap}>
        <span style={tag}>Border</span>
        <input
          type="color"
          title="Border color"
          value={strokeHex}
          onMouseDown={(e) => e.preventDefault()}
          onChange={(e) => setStroke({ color: { kind: 'rgb', hex: e.target.value } })}
          style={{ ...swatch, opacity: stroke ? 1 : 0.4 }}
        />
      </span>
      <select
        title="Border width"
        value={stroke ? String(stroke.width) : '0'}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => (e.target.value === '0' ? noBorder() : setStroke({ width: Number(e.target.value) }))}
        style={field}
      >
        <option value="0">None</option>
        {WIDTHS.map((w) => (
          <option key={w} value={w}>{w}px</option>
        ))}
      </select>
      <select
        title="Border style"
        value={stroke?.dash ?? 'solid'}
        disabled={!stroke}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => setStroke({ dash: e.target.value as StrokeDash })}
        style={{ ...field, opacity: stroke ? 1 : 0.5 }}
      >
        {DASHES.map((d) => (
          <option key={d.v} value={d.v}>{d.label}</option>
        ))}
      </select>
    </>
  );
}
