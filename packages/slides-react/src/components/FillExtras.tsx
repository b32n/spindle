// FillExtras — gradient presets and image fill for the fill controls (shape, table cell, slide background). The solid
// colour input stays with each caller. Images come from a file picker only and are stored as a data URL: nothing is
// ever fetched from a URL.

import React, { useRef } from 'react';
import { GRADIENT_PRESETS, isSafePictureSrc, type Fill } from '@b32nio/spindle-slides-core';

const field: React.CSSProperties = {
  height: 30, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 8,
  background: 'rgba(255,255,255,0.8)', color: '#334155', padding: '0 6px',
  fontSize: 13, fontFamily: '"Inter", sans-serif', cursor: 'pointer',
};

/** Read an image file as a data URL; resolves `null` for anything that is not a safe image data URL. */
export function readImageAsFill(file: File): Promise<Fill | null> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onerror = () => resolve(null);
    reader.onload = () => {
      const src = typeof reader.result === 'string' ? reader.result : '';
      resolve(isSafePictureSrc(src) ? { kind: 'picture', src, mode: 'stretch' } : null);
    };
    reader.readAsDataURL(file);
  });
}

export function FillExtras({
  onPick,
  disabled = false,
  testId = 'fill',
  style,
}: {
  onPick: (fill: Fill) => void;
  disabled?: boolean;
  testId?: string;
  style?: React.CSSProperties;
}): React.ReactElement {
  const fileRef = useRef<HTMLInputElement>(null);
  const onFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    const fill = file ? await readImageAsFill(file) : null;
    if (fill) onPick(fill);
  };
  return (
    <>
      <select
        title="Gradient fill"
        data-testid={`${testId}-gradient`}
        disabled={disabled}
        value=""
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => {
          const preset = GRADIENT_PRESETS.find((p) => p.id === e.target.value);
          if (preset) onPick(preset.fill);
        }}
        style={{ ...field, ...style, opacity: disabled ? 0.5 : 1 }}
      >
        <option value="">Gradient…</option>
        {GRADIENT_PRESETS.map((p) => (
          <option key={p.id} value={p.id}>{p.name}</option>
        ))}
      </select>
      <button
        type="button"
        title="Image fill"
        data-testid={`${testId}-image`}
        disabled={disabled}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => fileRef.current?.click()}
        style={{ ...field, ...style, opacity: disabled ? 0.5 : 1 }}
      >
        Image…
      </button>
      <input ref={fileRef} type="file" accept="image/*" onChange={onFile} style={{ display: 'none' }} />
    </>
  );
}
