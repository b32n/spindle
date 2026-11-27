import React from 'react';
import { opaqueLabel, type OpaqueElement } from '@b32nio/spindle-slides-core';

/**
 * Content from an imported PPTX the editor keeps but cannot draw (chart, SmartArt, OLE object, clip, ink): a labelled
 * placeholder ("Chart: <title> (kept, not editable)") in the element's box, with the clip's poster or the object's
 * preview behind the label when the source had one. It moves, resizes and deletes like any element.
 */
export function OpaqueView({ el }: { el: OpaqueElement }): React.ReactElement {
  const label = opaqueLabel(el);
  return (
    <div
      role="img"
      aria-label={el.alt ? `${label}. ${el.alt}` : label}
      data-opaque-kind={el.kind}
      style={{
        position: 'relative',
        width: '100%',
        height: '100%',
        boxSizing: 'border-box',
        overflow: 'hidden',
        border: '1px dashed #9ca3af',
        background: '#f3f4f6',
        color: '#4b5563',
        font: '14px/1.3 system-ui, sans-serif',
        userSelect: 'none',
      }}
    >
      {el.poster ? <img src={el.poster} alt="" draggable={false} style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', objectFit: 'contain', opacity: 0.45 }} /> : null}
      <div style={{ position: 'relative', display: 'flex', alignItems: 'center', justifyContent: 'center', width: '100%', height: '100%', padding: 8, textAlign: 'center' }}>{label}</div>
    </div>
  );
}
