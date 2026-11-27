// RunFields — the "Run options" section of the text options panel: baseline (superscript/subscript), caps, character
// spacing and language of the selected text, and inserting a slide-number or date field. Pure like ParagraphFields: the
// caller supplies the current run values and applies each change as a mark (null removes it).

import React, { useState } from 'react';
import { safeLang, type DateFieldFormat, type FieldKind, type ResolvedRun } from '@b32nio/spindle-slides-core';
import { DraftNumber } from './TextBodyFields';

export const SUPERSCRIPT_PCT = 30;
export const SUBSCRIPT_PCT = -25;

export interface RunMarkPatch {
  type: 'baseline' | 'caps' | 'charSpacing' | 'lang';
  /** The mark's attrs, or null to remove the mark. */
  attrs: Record<string, unknown> | null;
}

export interface FieldInsert {
  kind: FieldKind;
  format?: DateFieldFormat;
}

export const FIELD_CHOICES: Array<{ v: string; label: string; insert: FieldInsert }> = [
  { v: 'slidenum', label: 'Slide number', insert: { kind: 'slidenum' } },
  { v: 'date', label: 'Date', insert: { kind: 'datetime', format: 'date' } },
  { v: 'time', label: 'Time', insert: { kind: 'datetime', format: 'time' } },
  { v: 'datetime', label: 'Date and time', insert: { kind: 'datetime', format: 'datetime' } },
];

const field: React.CSSProperties = { height: 26, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 6, padding: '0 6px', fontSize: 13 };

export function RunFields({
  run,
  onMark,
  onInsertField,
  idPrefix = 'run',
}: {
  run: ResolvedRun;
  onMark: (patch: RunMarkPatch) => void;
  /** Absent when no field can be inserted (the text is not being edited). */
  onInsertField?: (insert: FieldInsert) => void;
  idPrefix?: string;
}): React.ReactElement {
  const id = (k: string) => `${idPrefix}-${k}`;
  const [langDraft, setLangDraft] = useState<string | null>(null);
  const pct = run.baselinePct;
  const baseline = pct === null ? 'normal' : pct === SUPERSCRIPT_PCT ? 'super' : pct === SUBSCRIPT_PCT ? 'sub' : 'custom';
  // Commits what was typed; with nothing typed there is nothing to commit (Enter on an untouched field must not clear it).
  const commitLang = () => {
    if (langDraft === null) return;
    const v = langDraft.trim();
    setLangDraft(null);
    if (v === '') onMark({ type: 'lang', attrs: null });
    else if (safeLang(v)) onMark({ type: 'lang', attrs: { lang: v } });
  };
  return (
    <div role="group" aria-label="Run options" style={{ display: 'grid', gridTemplateColumns: 'auto 150px', gap: 6, alignItems: 'center', fontSize: 13, color: '#334155', marginTop: 10, paddingTop: 10, borderTop: '1px solid rgba(0,0,0,0.08)' }}>
      <label htmlFor={id('baseline')}>Baseline</label>
      <select
        id={id('baseline')}
        data-testid="run-baseline"
        value={baseline}
        onChange={(e) => {
          const v = e.target.value;
          if (v === 'normal') onMark({ type: 'baseline', attrs: null });
          else if (v !== 'custom') onMark({ type: 'baseline', attrs: { pct: v === 'super' ? SUPERSCRIPT_PCT : SUBSCRIPT_PCT } });
        }}
        style={field}
      >
        <option value="normal">Normal</option>
        <option value="super">Superscript</option>
        <option value="sub">Subscript</option>
        {baseline === 'custom' && <option value="custom">{`Raised ${pct}%`}</option>}
      </select>
      <label htmlFor={id('caps')}>Capitals</label>
      <select
        id={id('caps')}
        data-testid="run-caps"
        value={run.caps ?? 'none'}
        onChange={(e) => onMark({ type: 'caps', attrs: e.target.value === 'none' ? null : { kind: e.target.value } })}
        style={field}
      >
        <option value="none">Normal</option>
        <option value="all">All caps</option>
        <option value="small">Small caps</option>
      </select>
      <label htmlFor={id('spacing')}>Character spacing (px)</label>
      <DraftNumber
        id={id('spacing')}
        step={0.5}
        data-testid="run-spacing"
        value={Math.round((run.spacingPx ?? 0) * 100) / 100}
        onValue={(n) => onMark({ type: 'charSpacing', attrs: n === 0 ? null : { px: n } })}
        style={field}
      />
      <label htmlFor={id('lang')}>Language (e.g. fr-FR)</label>
      <input
        id={id('lang')}
        data-testid="run-lang"
        type="text"
        maxLength={20}
        value={langDraft ?? run.lang ?? ''}
        onChange={(e) => setLangDraft(e.target.value)}
        onBlur={commitLang}
        onKeyDown={(e) => e.key === 'Enter' && commitLang()}
        aria-invalid={langDraft !== null && langDraft.trim() !== '' && !safeLang(langDraft.trim())}
        style={field}
      />
      <label htmlFor={id('field')}>Insert field</label>
      <select
        id={id('field')}
        data-testid="run-insert-field"
        value=""
        disabled={!onInsertField}
        title={onInsertField ? undefined : 'Edit the text to insert a field'}
        onChange={(e) => {
          const choice = FIELD_CHOICES.find((c) => c.v === e.target.value);
          if (choice && onInsertField) onInsertField(choice.insert);
        }}
        style={field}
      >
        <option value="">Choose…</option>
        {FIELD_CHOICES.map((c) => (
          <option key={c.v} value={c.v}>{c.label}</option>
        ))}
      </select>
    </div>
  );
}
