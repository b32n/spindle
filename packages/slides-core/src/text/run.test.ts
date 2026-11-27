import {
  baselineCss,
  fieldMarks,
  initialFieldText,
  resolveFieldText,
  resolveRun,
  runCss,
  safeFieldKind,
  safeLang,
  BASELINE_SCALE,
  MAX_BASELINE_PCT,
  MAX_SPACING_PX,
  SLIDENUM_PLACEHOLDER,
  type FieldContext,
} from './run';

const mark = (type: string, attrs?: Record<string, unknown>) => ({ type, attrs });

describe('resolveRun', () => {
  it('is empty without marks', () => {
    const none = { baselinePct: null, caps: null, spacingPx: null, lang: null };
    expect(resolveRun(undefined)).toEqual(none);
    expect(resolveRun([])).toEqual(none);
    expect(resolveRun([mark('bold')])).toEqual(none);
  });

  it('reads all four run marks', () => {
    expect(
      resolveRun([mark('baseline', { pct: 30 }), mark('caps', { kind: 'small' }), mark('charSpacing', { px: 1.5 }), mark('lang', { lang: 'fr-FR' })])
    ).toEqual({ baselinePct: 30, caps: 'small', spacingPx: 1.5, lang: 'fr-FR' });
  });

  it('treats zero, non-numeric and missing values as absent', () => {
    expect(resolveRun([mark('baseline', { pct: 0 }), mark('charSpacing', { px: 0 })])).toMatchObject({ baselinePct: null, spacingPx: null });
    expect(resolveRun([mark('baseline', { pct: 'up' }), mark('charSpacing', {})])).toMatchObject({ baselinePct: null, spacingPx: null });
    expect(resolveRun([mark('baseline'), mark('caps'), mark('charSpacing'), mark('lang')])).toEqual({ baselinePct: null, caps: null, spacingPx: null, lang: null });
  });

  it('accepts numeric strings and clamps extreme values', () => {
    expect(resolveRun([mark('baseline', { pct: '25' })]).baselinePct).toBe(25);
    expect(resolveRun([mark('baseline', { pct: 9000 })]).baselinePct).toBe(MAX_BASELINE_PCT);
    expect(resolveRun([mark('baseline', { pct: -9000 })]).baselinePct).toBe(-MAX_BASELINE_PCT);
    expect(resolveRun([mark('charSpacing', { px: 1e9 })]).spacingPx).toBe(MAX_SPACING_PX);
    expect(resolveRun([mark('charSpacing', { px: -1e9 })]).spacingPx).toBe(-MAX_SPACING_PX);
  });

  it('allowlists the caps kind', () => {
    expect(resolveRun([mark('caps', { kind: 'all' })]).caps).toBe('all');
    expect(resolveRun([mark('caps', { kind: 'all;background:url(//evil)' })]).caps).toBeNull();
  });

  it('the last mark of a type wins', () => {
    expect(resolveRun([mark('caps', { kind: 'all' }), mark('caps', { kind: 'small' })]).caps).toBe('small');
  });
});

describe('safeLang', () => {
  it.each(['en', 'en-US', 'zh-Hans-CN', 'de-DE-1996', 'fil', 'es-419'])('accepts %s', (v) => expect(safeLang(v)).toBe(v));
  it.each(['', 'e', 'en_US', 'en-', 'x"onload=1', 'en-US;color:red', 'abcd', 'en-US-aaaaaaaaa', 5, null, undefined, {}])('rejects %p', (v) => expect(safeLang(v)).toBeNull());
});

describe('css helpers', () => {
  it('runCss maps caps and spacing', () => {
    const base = { baselinePct: null, caps: null, spacingPx: null, lang: null };
    expect(runCss(base)).toEqual({});
    expect(runCss({ ...base, caps: 'all', spacingPx: 2 })).toEqual({ textTransform: 'uppercase', letterSpacing: '2px' });
    expect(runCss({ ...base, caps: 'small', spacingPx: -1 })).toEqual({ fontVariant: 'small-caps', letterSpacing: '-1px' });
  });

  it('baselineCss shifts up for a raise and down for a lowering, in em of the smaller size', () => {
    const up = baselineCss(30);
    expect(up).toMatchObject({ position: 'relative', fontSize: `${BASELINE_SCALE}em`, lineHeight: 0 });
    expect(up.top).toBe('-0.462em');
    expect(baselineCss(-25).top).toBe('0.385em');
  });
});

describe('fields', () => {
  const now = new Date(2026, 0, 15, 14, 30);
  const ctx = (patch: Partial<FieldContext> = {}): FieldContext => ({ slideNumber: 4, now, locale: 'en-US', ...patch });

  it('slide number shows the number, else the stored text, else the placeholder', () => {
    expect(resolveFieldText({ kind: 'slidenum' }, ctx(), '‹#›')).toBe('4');
    expect(resolveFieldText({ kind: 'slidenum' }, ctx({ slideNumber: null }), '9')).toBe('9');
    expect(resolveFieldText({ kind: 'slidenum' }, ctx({ slideNumber: null }))).toBe(SLIDENUM_PLACEHOLDER);
    expect(resolveFieldText({ kind: 'slidenum' }, ctx({ slideNumber: NaN }))).toBe(SLIDENUM_PLACEHOLDER);
  });

  it('date and time follow the format and locale', () => {
    expect(resolveFieldText({ kind: 'datetime', format: 'date' }, ctx())).toBe('Jan 15, 2026');
    expect(resolveFieldText({ kind: 'datetime' }, ctx())).toBe('Jan 15, 2026');
    expect(resolveFieldText({ kind: 'datetime', format: null }, ctx())).toBe('Jan 15, 2026');
    expect(resolveFieldText({ kind: 'datetime', format: 'time' }, ctx())).toMatch(/2:30\sPM/);
    expect(resolveFieldText({ kind: 'datetime', format: 'datetime' }, ctx())).toMatch(/Jan 15, 2026.*2:30\sPM/);
    expect(resolveFieldText({ kind: 'datetime', format: 'date' }, ctx({ locale: 'de-DE' }))).toBe('15.01.2026');
  });

  it('a date that cannot be formatted falls back to the stored text', () => {
    expect(resolveFieldText({ kind: 'datetime' }, ctx({ now: new Date(NaN) }), '1/15/26')).toBe('1/15/26');
  });

  it('a crafted format shows a plain date; footer shows its stored text', () => {
    expect(resolveFieldText({ kind: 'datetime', format: 'x' as never }, ctx())).toBe('Jan 15, 2026');
    expect(resolveFieldText({ kind: 'footer' }, ctx(), 'Acme Inc')).toBe('Acme Inc');
    expect(resolveFieldText({ kind: 'footer' }, ctx())).toBe('');
    expect(resolveFieldText(undefined, ctx(), 'kept')).toBe('kept');
  });

  it('safeFieldKind degrades unknown kinds to footer', () => {
    expect(safeFieldKind('slidenum')).toBe('slidenum');
    expect(safeFieldKind('datetime')).toBe('datetime');
    expect(safeFieldKind('evil')).toBe('footer');
    expect(safeFieldKind(undefined)).toBe('footer');
  });

  it('fieldMarks reads the marks of the cached text, then of the node', () => {
    const field = (patch: object) => ({ type: 'field' as const, attrs: { kind: 'footer' as const }, ...patch });
    expect(fieldMarks(field({}))).toEqual([]);
    expect(fieldMarks(field({ content: [{ type: 'text', text: 'x', marks: [{ type: 'bold' }] }], marks: [{ type: 'italic' }] }))).toEqual([{ type: 'bold' }, { type: 'italic' }]);
    expect(fieldMarks(field({ content: [{ type: 'text', text: 'x' }] }))).toEqual([]);
  });

  it('initialFieldText is what the field shows now', () => {
    expect(initialFieldText({ kind: 'slidenum' })).toBe(SLIDENUM_PLACEHOLDER);
    expect(initialFieldText({ kind: 'footer' })).toBe('');
    expect(initialFieldText({ kind: 'datetime', format: 'date' }, new Date(2026, 5, 1))).toMatch(/2026/);
    expect(initialFieldText({ kind: 'datetime' })).toMatch(/\d{4}/);
  });
});
