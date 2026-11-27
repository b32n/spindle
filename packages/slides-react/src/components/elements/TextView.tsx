import React from 'react';
import {
  autofitMode,
  resolveColor,
  resolveFont,
  isRichTextEmpty,
  type TextElement,
  type ThemeData,
} from '@b32nio/spindle-slides-core';
import { StaticRichText } from './StaticRichText';
import { RichTextEditor } from '../RichTextEditor';
import { useDeck, useEditingId } from '../../hooks';
import { strokeAttrs } from './style';
import { htmlFill } from './fill';
import { bodyBoxStyle, bodyOverflow, columnsStyle, scaleVars } from './body';
import { textBodyRef } from '../../text-fit';

/** Element-level text defaults (PPTX defRPr analogue) applied to the container. */
export function bodyDefaults(body: TextElement['bodyStyle'], theme: ThemeData): React.CSSProperties {
  const b = body ?? {};
  return {
    fontFamily: resolveFont(b.fontFamily, theme),
    // Under shrink-to-fit the size follows the body's font scale variable, so measuring can rescale it live.
    fontSize: autofitMode(b) === 'shrink' ? `calc(${b.fontSize ?? 18}px * var(--spindle-fs, 1))` : (b.fontSize ?? 18),
    fontWeight: b.bold ? 700 : 400,
    color: b.color ? resolveColor(b.color, theme) : resolveColor({ kind: 'theme', slot: 'dk1' }, theme),
  };
}

export function TextView({ el, theme, interactive = false }: { el: TextElement; theme: ThemeData; interactive?: boolean }): React.ReactElement {
  const deck = useDeck();
  const body = el.bodyStyle ?? {};
  const { style: fillStyle, layer: fillLayer } = htmlFill(el.fill, theme);
  const border = strokeAttrs(el.stroke, theme);
  // Only the interactive stage mounts the live editor — never the filmstrip
  // thumbnails or a read-only view (two editors on one element fight over focus).
  const editingId = useEditingId();
  const editing = interactive && editingId === el.id;
  const empty = isRichTextEmpty(el.richText);
  const prompt = empty && el.placeholder ? deck.getPlaceholderPrompt(el.containerId, el.placeholder) : undefined;

  const scaled = autofitMode(body) === 'shrink';
  const innerRef = React.useMemo(() => textBodyRef(el.id, interactive), [el.id, interactive]);
  const cols = columnsStyle(body);

  const container: React.CSSProperties = {
    width: '100%',
    height: '100%',
    boxSizing: 'border-box',
    overflow: bodyOverflow(body),
    ...fillStyle,
    ...(border ? { border: `${border.strokeWidth}px solid ${border.stroke}` } : {}),
  };

  const inner: React.CSSProperties = {
    width: '100%',
    height: '100%',
    ...bodyBoxStyle(body),
    ...(scaled ? scaleVars(body) : {}),
    whiteSpace: body.wrap === false ? 'nowrap' : 'normal',
    ...bodyDefaults(body, theme),
  };

  return (
    <div style={container}>
      {fillLayer}
      {editing ? (
        <RichTextEditor elementId={el.id} theme={theme} bodyStyle={body} />
      ) : (
        <div ref={innerRef} data-text-body style={inner}>
          {empty
            ? prompt
              ? <div style={{ opacity: 0.4, textAlign: el.richText.content[0]?.attrs?.align ?? 'left' }}>{prompt}</div>
              : null
            : cols
              ? <div style={cols}><StaticRichText doc={el.richText} theme={theme} scaled={scaled} /></div>
              : <StaticRichText doc={el.richText} theme={theme} scaled={scaled} />}
        </div>
      )}
    </div>
  );
}
