import React from 'react';
import { resolveFill, type ShapeElement, type ThemeData } from '@b32nio/spindle-slides-core';
import { shapeGeom } from './shapes';
import { StaticRichText } from './StaticRichText';
import { RichTextEditor } from '../RichTextEditor';
import { useEditingId } from '../../hooks';
import { strokeAttrs } from './style';

export function ShapeView({ el, theme, interactive = false }: { el: ShapeElement; theme: ThemeData; interactive?: boolean }): React.ReactElement {
  const editingId = useEditingId();
  const editing = interactive && editingId === el.id;
  const w = Math.max(1, el.w);
  const h = Math.max(1, el.h);
  const geom = shapeGeom(el.shape, w, h, el.adjustments);
  const fill = resolveFill(el.fill, theme) ?? 'none';
  const stroke = strokeAttrs(el.stroke, theme);

  const flip = `${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1}`;
  const svgStyle: React.CSSProperties = {
    display: 'block',
    width: '100%',
    height: '100%',
    transform: el.flipH || el.flipV ? `scale(${flip})` : undefined,
    transformOrigin: 'center center',
  };

  const geomProps = {
    fill,
    ...(stroke ? { stroke: stroke.stroke, strokeWidth: stroke.strokeWidth, strokeDasharray: stroke.strokeDasharray } : {}),
  };

  return (
    <div style={{ position: 'relative', width: '100%', height: '100%' }}>
      <svg viewBox={`0 0 ${w} ${h}`} preserveAspectRatio="none" style={svgStyle}>
        {geom.type === 'ellipse' ? (
          <ellipse cx={geom.cx} cy={geom.cy} rx={geom.rx} ry={geom.ry} {...geomProps} />
        ) : (
          <path d={geom.d} {...geomProps} />
        )}
      </svg>
      {editing ? (
        <div style={{ position: 'absolute', inset: 0 }}>
          <RichTextEditor elementId={el.id} theme={theme} bodyStyle={el.bodyStyle} centered />
        </div>
      ) : el.richText ? (
        <div
          style={{
            position: 'absolute',
            inset: 0,
            display: 'flex',
            flexDirection: 'column',
            justifyContent: el.bodyStyle?.vAlign === 'top' ? 'flex-start' : el.bodyStyle?.vAlign === 'bottom' ? 'flex-end' : 'center',
            padding: el.bodyStyle?.padding ?? 8,
            boxSizing: 'border-box',
            fontFamily: theme.fonts.minor,
            fontSize: 18,
            textAlign: 'center',
            overflow: 'hidden',
          }}
        >
          <StaticRichText doc={el.richText} theme={theme} />
        </div>
      ) : null}
    </div>
  );
}
