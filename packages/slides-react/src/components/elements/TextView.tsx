import React from 'react';
import {
  resolveColor,
  resolveFill,
  resolveFont,
  isRichTextEmpty,
  type TextElement,
  type ThemeData,
} from '@b32nio/spindle-slides-core';
import { StaticRichText } from './StaticRichText';
import { RichTextEditor } from '../RichTextEditor';
import { useDeck, useEditingId } from '../../hooks';
import { strokeAttrs } from './style';

const V_ALIGN: Record<NonNullable<TextElement['bodyStyle']>['vAlign'] & string, string> = {
  top: 'flex-start',
  middle: 'center',
  bottom: 'flex-end',
};

/** Element-level text defaults (PPTX defRPr analogue) applied to the container. */
export function bodyDefaults(body: TextElement['bodyStyle'], theme: ThemeData): React.CSSProperties {
  const b = body ?? {};
  return {
    fontFamily: resolveFont(b.fontFamily, theme),
    fontSize: b.fontSize ?? 18,
    fontWeight: b.bold ? 700 : 400,
    color: b.color ? resolveColor(b.color, theme) : resolveColor({ kind: 'theme', slot: 'dk1' }, theme),
  };
}

export function TextView({ el, theme, interactive = false }: { el: TextElement; theme: ThemeData; interactive?: boolean }): React.ReactElement {
  const deck = useDeck();
  const body = el.bodyStyle ?? {};
  const bg = el.fill ? resolveFill(el.fill, theme) : null;
  const border = strokeAttrs(el.stroke, theme);
  // Only the interactive stage mounts the live editor — never the filmstrip
  // thumbnails or a read-only view (two editors on one element fight over focus).
  const editingId = useEditingId();
  const editing = interactive && editingId === el.id;
  const empty = isRichTextEmpty(el.richText);
  const prompt = empty && el.placeholder ? deck.getPlaceholderPrompt(el.containerId, el.placeholder) : undefined;

  const container: React.CSSProperties = {
    width: '100%',
    height: '100%',
    boxSizing: 'border-box',
    overflow: 'hidden',
    ...(bg ? { background: bg } : {}),
    ...(border ? { border: `${border.strokeWidth}px solid ${border.stroke}` } : {}),
  };

  const inner: React.CSSProperties = {
    width: '100%',
    height: '100%',
    boxSizing: 'border-box',
    padding: body.padding ?? 8,
    display: 'flex',
    flexDirection: 'column',
    justifyContent: V_ALIGN[body.vAlign ?? 'top'],
    whiteSpace: body.wrap === false ? 'nowrap' : 'normal',
    ...bodyDefaults(body, theme),
  };

  return (
    <div style={container}>
      {editing ? (
        <RichTextEditor elementId={el.id} theme={theme} bodyStyle={body} />
      ) : (
        <div style={inner}>
          {empty
            ? prompt
              ? <div style={{ opacity: 0.4, textAlign: el.richText.content[0]?.attrs?.align ?? 'left' }}>{prompt}</div>
              : null
            : <StaticRichText doc={el.richText} theme={theme} />}
        </div>
      )}
    </div>
  );
}
