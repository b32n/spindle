// TextBodyFields — the "Text box options" panel body: how the text sits in its box (insets, centred block, direction,
// columns) and what happens when it does not fit (autofit). Pure (no hooks) so the controls are unit-testable; the
// caller supplies the current body style and applies each change.

import React from 'react';
import { autofitMode, resolveColumns, resolveInsets, textVert, type BodyInsets, type BodyStyle, type TextVert } from '@b32nio/spindle-slides-core';

export const AUTOFIT_OPTIONS: Array<{ v: 'none' | 'shrink' | 'resize'; label: string }> = [
  { v: 'none', label: 'Do not autofit' },
  { v: 'shrink', label: 'Shrink text on overflow' },
  { v: 'resize', label: 'Resize shape to fit text' },
];

export const DIRECTION_OPTIONS: Array<{ v: TextVert; label: string }> = [
  { v: 'horz', label: 'Horizontal' },
  { v: 'vert', label: 'Rotate all text 90°' },
  { v: 'eaVert', label: 'Stacked (East Asian)' },
];

const SIDES: Array<{ k: keyof BodyInsets; label: string }> = [
  { k: 'l', label: 'Left' },
  { k: 't', label: 'Top' },
  { k: 'r', label: 'Right' },
  { k: 'b', label: 'Bottom' },
];

const field: React.CSSProperties = { height: 26, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 6, padding: '0 6px', fontSize: 13 };

export function TextBodyFields({
  body,
  onChange,
  idPrefix = 'text-body',
}: {
  body: BodyStyle | undefined;
  onChange: (patch: Partial<BodyStyle>) => void;
  idPrefix?: string;
}): React.ReactElement {
  const insets = resolveInsets(body);
  const cols = resolveColumns(body);
  return (
    <div role="group" aria-label="Text box options" style={{ display: 'grid', gridTemplateColumns: 'auto 150px', gap: 6, alignItems: 'center', fontSize: 13, color: '#334155' }}>
      <label htmlFor={`${idPrefix}-autofit`}>Autofit</label>
      <select
        id={`${idPrefix}-autofit`}
        data-testid="text-autofit"
        value={autofitMode(body)}
        onChange={(e) => onChange({ autofit: { mode: e.target.value as 'none' | 'shrink' | 'resize' } })}
        style={field}
      >
        {AUTOFIT_OPTIONS.map((o) => (
          <option key={o.v} value={o.v}>{o.label}</option>
        ))}
      </select>
      <label htmlFor={`${idPrefix}-vert`}>Direction</label>
      <select
        id={`${idPrefix}-vert`}
        data-testid="text-vert"
        value={textVert(body)}
        onChange={(e) => onChange({ vert: e.target.value === 'horz' ? undefined : (e.target.value as TextVert) })}
        style={field}
      >
        {DIRECTION_OPTIONS.map((o) => (
          <option key={o.v} value={o.v}>{o.label}</option>
        ))}
      </select>
      {SIDES.map(({ k, label }) => (
        <React.Fragment key={k}>
          <label htmlFor={`${idPrefix}-inset-${k}`}>{label} inset (px)</label>
          <input
            id={`${idPrefix}-inset-${k}`}
            type="number"
            min={0}
            step={1}
            data-testid={`text-inset-${k}`}
            value={Math.round(insets[k] * 100) / 100}
            onChange={(e) => {
              const v = e.target.valueAsNumber;
              onChange({ insets: { ...insets, [k]: Number.isFinite(v) ? Math.max(0, v) : 0 } });
            }}
            style={field}
          />
        </React.Fragment>
      ))}
      <label htmlFor={`${idPrefix}-cols`}>Columns</label>
      <input
        id={`${idPrefix}-cols`}
        type="number"
        min={1}
        max={16}
        step={1}
        data-testid="text-columns"
        value={cols.count}
        onChange={(e) => {
          const n = e.target.valueAsNumber;
          onChange({ columns: Number.isFinite(n) && n > 1 ? { count: n, gap: cols.gap } : undefined });
        }}
        style={field}
      />
      <label htmlFor={`${idPrefix}-colgap`}>Column spacing (px)</label>
      <input
        id={`${idPrefix}-colgap`}
        type="number"
        min={0}
        step={1}
        data-testid="text-column-gap"
        value={cols.gap}
        disabled={cols.count < 2}
        onChange={(e) => {
          const g = e.target.valueAsNumber;
          onChange({ columns: { count: cols.count, gap: Number.isFinite(g) ? Math.max(0, g) : 0 } });
        }}
        style={field}
      />
      <label htmlFor={`${idPrefix}-ctr`}>Centre text block</label>
      <input
        id={`${idPrefix}-ctr`}
        type="checkbox"
        data-testid="text-anchor-ctr"
        checked={!!body?.anchorCtr}
        onChange={(e) => onChange({ anchorCtr: e.target.checked ? true : undefined })}
      />
    </div>
  );
}
