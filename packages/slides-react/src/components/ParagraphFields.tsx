// ParagraphFields — the "Paragraph options" panel body: bullet (kind, glyph, numbering, colour, size), line height,
// spacing and margins of the current paragraph. Pure (no hooks) like TextBodyFields: the caller supplies the paragraph's
// attrs and applies each change as a paragraph patch (null clears an attr).

import React from 'react';
import { resolveParagraph, type BulletScheme, type ParagraphAttrs, type Bullet, type Color } from '@b32nio/spindle-slides-core';
import { DraftNumber } from './TextBodyFields';

export const BULLET_CHOICES: Array<{ v: 'default' | 'none' | 'char' | 'auto'; label: string }> = [
  { v: 'default', label: 'Per list setting' },
  { v: 'none', label: 'No bullet' },
  { v: 'char', label: 'Bullet character' },
  { v: 'auto', label: 'Numbered' },
];

export const SCHEME_CHOICES: Array<{ v: BulletScheme; label: string }> = [
  { v: 'arabic', label: '1, 2, 3' },
  { v: 'alphaLower', label: 'a, b, c' },
  { v: 'alphaUpper', label: 'A, B, C' },
  { v: 'romanLower', label: 'i, ii, iii' },
  { v: 'romanUpper', label: 'I, II, III' },
];

const field: React.CSSProperties = { height: 26, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 6, padding: '0 6px', fontSize: 13 };

/** The #rrggbb value a colour input can show for a bullet colour (theme colours show black until picked). */
function colourInputValue(c: Color | undefined): string {
  return c && c.kind === 'rgb' && /^#[0-9a-fA-F]{6}$/.test(c.hex) ? c.hex : '#000000';
}

const px = (n: number | null | undefined): number => Math.round((n ?? 0) * 100) / 100;

export function ParagraphFields({
  attrs,
  onChange,
  idPrefix = 'para',
}: {
  attrs: ParagraphAttrs | undefined;
  onChange: (patch: Partial<ParagraphAttrs>) => void;
  idPrefix?: string;
}): React.ReactElement {
  const a = attrs ?? {};
  const r = resolveParagraph(a);
  const bullet = a.bullet && typeof a.bullet === 'object' ? a.bullet : undefined;
  const kind = bullet ? bullet.kind : 'default';
  const setBullet = (patch: Partial<Bullet>) => onChange({ bullet: { kind: 'char', ...bullet, ...patch } });
  const exact = 'px' in r.lineHeight;
  const spaceBefore = r.spaceBefore && 'px' in r.spaceBefore ? r.spaceBefore.px : 0;
  const spaceAfter = r.spaceAfter && 'px' in r.spaceAfter ? r.spaceAfter.px : 0;
  const id = (k: string) => `${idPrefix}-${k}`;
  return (
    <div role="group" aria-label="Paragraph options" style={{ display: 'grid', gridTemplateColumns: 'auto 150px', gap: 6, alignItems: 'center', fontSize: 13, color: '#334155', marginTop: 10, paddingTop: 10, borderTop: '1px solid rgba(0,0,0,0.08)' }}>
      <label htmlFor={id('bullet')}>Bullet</label>
      <select
        id={id('bullet')}
        data-testid="para-bullet"
        value={kind}
        onChange={(e) => {
          const v = e.target.value;
          if (v === 'default') onChange({ bullet: null });
          else if (v === 'none') onChange({ bullet: { kind: 'none' } });
          else onChange({ bullet: { ...bullet, kind: v as 'char' | 'auto' } });
        }}
        style={field}
      >
        {BULLET_CHOICES.map((o) => (
          <option key={o.v} value={o.v}>{o.label}</option>
        ))}
      </select>
      {kind === 'char' && (
        <>
          <label htmlFor={id('char')}>Character</label>
          <input
            id={id('char')}
            data-testid="para-bullet-char"
            type="text"
            maxLength={4}
            value={bullet?.char ?? '•'}
            onChange={(e) => setBullet({ kind: 'char', char: e.target.value || undefined })}
            style={field}
          />
        </>
      )}
      {kind === 'auto' && (
        <>
          <label htmlFor={id('scheme')}>Numbering</label>
          <select
            id={id('scheme')}
            data-testid="para-bullet-scheme"
            value={r.marker!.scheme}
            onChange={(e) => setBullet({ kind: 'auto', scheme: e.target.value as BulletScheme })}
            style={field}
          >
            {SCHEME_CHOICES.map((o) => (
              <option key={o.v} value={o.v}>{o.label}</option>
            ))}
          </select>
          <label htmlFor={id('start')}>Start at</label>
          <DraftNumber
            id={id('start')}
            min={1}
            step={1}
            data-testid="para-bullet-start"
            value={r.marker!.startAt ?? 1}
            onValue={(n) => setBullet({ kind: 'auto', startAt: Math.max(1, Math.trunc(n)) })}
            style={field}
          />
        </>
      )}
      {(kind === 'char' || kind === 'auto') && (
        <>
          <label htmlFor={id('bullet-size')}>Bullet size (%)</label>
          <DraftNumber
            id={id('bullet-size')}
            min={25}
            max={400}
            step={5}
            data-testid="para-bullet-size"
            value={r.marker!.sizePct}
            onValue={(n) => setBullet({ sizePct: Math.min(400, Math.max(25, n)) })}
            style={field}
          />
          <label htmlFor={id('bullet-color')}>Bullet colour</label>
          <input
            id={id('bullet-color')}
            data-testid="para-bullet-color"
            type="color"
            value={colourInputValue(bullet?.color)}
            onChange={(e) => setBullet({ color: { kind: 'rgb', hex: e.target.value } })}
            style={{ ...field, width: 40, padding: 2 }}
          />
        </>
      )}
      <label htmlFor={id('lh-mode')}>Line height</label>
      <select
        id={id('lh-mode')}
        data-testid="para-lh-mode"
        value={exact ? 'exact' : 'multiple'}
        onChange={(e) => onChange(e.target.value === 'exact' ? { lineHeightPx: 24 } : { lineHeightPx: null })}
        style={field}
      >
        <option value="multiple">Multiple of text size</option>
        <option value="exact">Exact (px)</option>
      </select>
      <label htmlFor={id('lh')}>{exact ? 'Line height (px)' : 'Line height (×)'}</label>
      <DraftNumber
        id={id('lh')}
        min={exact ? 1 : 0.5}
        step={exact ? 1 : 0.05}
        data-testid="para-lh"
        value={'px' in r.lineHeight ? px(r.lineHeight.px) : px(r.lineHeight.mult)}
        onValue={(n) => (exact ? onChange({ lineHeightPx: Math.max(1, n) }) : onChange({ lineHeight: Math.max(0.5, n) }))}
        style={field}
      />
      <label htmlFor={id('before')}>Space before (px)</label>
      <DraftNumber
        id={id('before')}
        min={0}
        step={1}
        data-testid="para-space-before"
        value={px(spaceBefore)}
        onValue={(n) => onChange({ spaceBefore: Math.max(0, n), spaceBeforePct: null })}
        style={field}
      />
      <label htmlFor={id('after')}>Space after (px)</label>
      <DraftNumber
        id={id('after')}
        min={0}
        step={1}
        data-testid="para-space-after"
        value={px(spaceAfter)}
        onValue={(n) => onChange({ spaceAfter: Math.max(0, n), spaceAfterPct: null })}
        style={field}
      />
      <label htmlFor={id('margin')}>Left margin (px)</label>
      <DraftNumber
        id={id('margin')}
        min={0}
        step={1}
        data-testid="para-margin-left"
        value={px(a.marginLeft ?? r.paddingLeft)}
        onValue={(n) => onChange({ marginLeft: Math.max(0, n) })}
        style={field}
      />
      <label htmlFor={id('first')}>First line indent (px)</label>
      <DraftNumber
        id={id('first')}
        step={1}
        data-testid="para-first-indent"
        value={px(a.firstLineIndent)}
        onValue={(n) => onChange({ firstLineIndent: n })}
        style={field}
      />
    </div>
  );
}
