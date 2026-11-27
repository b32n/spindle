// ImageCropFields — the crop popover body: the percent trimmed from each edge of the picture, and a reset. Pure
// (no hooks) so the controls are unit-testable; the caller supplies the current crop and applies each change.

import React from 'react';
import { NO_CROP, normalizeCrop, setCropEdge, type ImageCrop } from '@b32nio/spindle-slides-core';

const EDGES: Array<{ k: keyof ImageCrop; label: string }> = [
  { k: 'l', label: 'Left' },
  { k: 't', label: 'Top' },
  { k: 'r', label: 'Right' },
  { k: 'b', label: 'Bottom' },
];

const field: React.CSSProperties = { height: 26, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 6, padding: '0 6px', fontSize: 13 };

export function ImageCropFields({
  crop,
  onChange,
  idPrefix = 'crop',
}: {
  idPrefix?: string;
  crop: ImageCrop | undefined;
  onChange: (crop: ImageCrop) => void;
}): React.ReactElement {
  const cur = { ...NO_CROP, ...normalizeCrop(crop) };
  return (
    <div role="group" aria-label="Crop image" style={{ display: 'grid', gridTemplateColumns: 'auto 64px', gap: 6, alignItems: 'center', fontSize: 13, color: '#334155' }}>
      {EDGES.map(({ k, label }) => (
        <React.Fragment key={k}>
          <label htmlFor={`${idPrefix}-${k}`}>{label} (%)</label>
          <input
            id={`${idPrefix}-${k}`}
            type="number"
            min={0}
            max={99}
            step={1}
            data-testid={`image-crop-${k}`}
            value={Math.round(cur[k] * 100) / 100}
            onChange={(e) => onChange(setCropEdge(crop, k, e.target.valueAsNumber))}
            style={field}
          />
        </React.Fragment>
      ))}
      <button
        type="button"
        aria-label="Reset crop"
        data-testid="image-crop-reset"
        onClick={() => onChange({ ...NO_CROP })}
        style={{ ...field, gridColumn: '1 / -1', height: 28, background: '#f8fafc', cursor: 'pointer' }}
      >
        Reset crop
      </button>
    </div>
  );
}
