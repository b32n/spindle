import React, { useId } from 'react';
import { autofitMode, normalizeCustomGeometry, type ShapeElement, type ThemeData } from '@b32nio/spindle-slides-core';
import { shapeGeom } from './shapes';
import { customPathD, pathShade } from './custom-geom';
import { StaticRichText } from './StaticRichText';
import { RichTextEditor } from '../RichTextEditor';
import { useEditingId } from '../../hooks';
import { strokeAttrs } from './style';
import { paintId, svgFill } from './fill';
import { bodyBoxStyle, bodyOverflow, columnsStyle, scaleVars } from './body';
import { textBodyRef } from '../../text-fit';

export function ShapeView({ el, theme, interactive = false }: { el: ShapeElement; theme: ThemeData; interactive?: boolean }): React.ReactElement {
  const instance = useId();
  const editingId = useEditingId();
  const editing = interactive && editingId === el.id;
  const w = Math.max(1, el.w);
  const h = Math.max(1, el.h);
  const geom = shapeGeom(el.shape, w, h, el.adjustments);
  const { defs, fill } = svgFill(el.fill, theme, paintId(el.id, instance), w, h);
  const stroke = strokeAttrs(el.stroke, theme);

  const shrink = autofitMode(el.bodyStyle) === 'shrink';
  const bodyRef = React.useMemo(() => textBodyRef(el.id, interactive), [el.id, interactive]);
  const cols = columnsStyle(el.bodyStyle);

  const flip = `${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1}`;
  const svgStyle: React.CSSProperties = {
    display: 'block',
    width: '100%',
    height: '100%',
    transform: el.flipH || el.flipV ? `scale(${flip})` : undefined,
    transformOrigin: 'center center',
  };

  // a `path` shape with no usable geometry falls back to the rectangle `shapeGeom` gives any non-preset name
  const custom = React.useMemo(() => (el.shape === 'path' ? normalizeCustomGeometry(el.path) : undefined), [el.shape, el.path]);

  const geomProps = {
    fill,
    ...(stroke ? { stroke: stroke.stroke, strokeWidth: stroke.strokeWidth, strokeDasharray: stroke.strokeDasharray } : {}),
  };

  return (
    <div style={{ position: 'relative', width: '100%', height: '100%' }}>
      <svg viewBox={`0 0 ${w} ${h}`} preserveAspectRatio="none" style={svgStyle}>
        {defs}
        {custom ? (
          custom.paths.map((p, i) => {
            const d = customPathD(p, w, h);
            const shade = pathShade(p.fill);
            return (
              <React.Fragment key={i}>
                <path d={d} fill={p.fill === 'none' ? 'none' : fill} stroke="none" />
                {shade ? <path d={d} fill={shade.color} fillOpacity={shade.opacity} stroke="none" /> : null}
                {p.stroke === false ? null : <path d={d} {...geomProps} fill="none" />}
              </React.Fragment>
            );
          })
        ) : geom.type === 'ellipse' ? (
          <ellipse cx={geom.cx} cy={geom.cy} rx={geom.rx} ry={geom.ry} {...geomProps} />
        ) : geom.strokeD ? (
          <>
            <path d={geom.d} fill={fill} stroke="none" />
            <path d={geom.strokeD} {...geomProps} fill="none" />
          </>
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
          ref={bodyRef}
          data-text-body
          style={{
            position: 'absolute',
            inset: 0,
            ...bodyBoxStyle(el.bodyStyle, 'middle'),
            ...(shrink ? scaleVars(el.bodyStyle) : {}),
            fontFamily: theme.fonts.minor,
            fontSize: shrink ? 'calc(18px * var(--spindle-fs, 1))' : 18,
            textAlign: 'center',
            overflow: bodyOverflow(el.bodyStyle),
          }}
        >
          {cols
            ? <div style={cols}><StaticRichText doc={el.richText} theme={theme} scaled={shrink} /></div>
            : <StaticRichText doc={el.richText} theme={theme} scaled={shrink} />}
        </div>
      ) : null}
    </div>
  );
}
