import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DeckImpl, getBuiltinTheme, slidesSchema, type RichTextDoc, type RichTextInline, type FieldContext, type ResolvedRun } from '@b32nio/spindle-slides-core';
import { DeckProvider } from '../context/DeckContext';
import { SlideView } from './SlideView';
import { StaticRichText, FieldContextValue } from './elements/StaticRichText';
import { RunFields, FIELD_CHOICES, SUPERSCRIPT_PCT, SUBSCRIPT_PCT, type FieldInsert } from './RunFields';
import { insertField } from '../interactions/formatting';

const theme = getBuiltinTheme();
const doc = (...content: RichTextInline[]): RichTextDoc => ({ type: 'doc', content: [{ type: 'paragraph', content }] });
const render = (d: RichTextDoc, ctx?: FieldContext) =>
  renderToStaticMarkup(
    <FieldContextValue.Provider value={ctx ?? null}>
      <StaticRichText doc={d} theme={theme} />
    </FieldContextValue.Provider>,
  );
const run = (text: string, marks: Array<{ type: string; attrs?: Record<string, unknown> }>): RichTextInline => ({ type: 'text', text, marks });

describe('StaticRichText run marks', () => {
  it('draws caps, character spacing and a language tag on the run', () => {
    expect(render(doc(run('a', [{ type: 'caps', attrs: { kind: 'all' } }])))).toContain('text-transform:uppercase');
    expect(render(doc(run('a', [{ type: 'caps', attrs: { kind: 'small' } }])))).toContain('font-variant:small-caps');
    expect(render(doc(run('a', [{ type: 'charSpacing', attrs: { px: 3 } }])))).toContain('letter-spacing:3px');
    expect(render(doc(run('a', [{ type: 'lang', attrs: { lang: 'fr-FR' } }])))).toContain('lang="fr-FR"');
  });

  it('draws a raised or lowered run as a smaller shifted inner span', () => {
    const up = render(doc(run('2', [{ type: 'baseline', attrs: { pct: 30 } }])));
    expect(up).toContain('top:-0.462em');
    expect(up).toContain('font-size:0.65em');
    expect(render(doc(run('2', [{ type: 'baseline', attrs: { pct: -25 } }])))).toContain('top:0.385em');
  });

  it('keeps a link link-shaped with a baseline and language', () => {
    const h = render(doc(run('x', [{ type: 'link', attrs: { href: 'https://e.com' } }, { type: 'baseline', attrs: { pct: 30 } }, { type: 'lang', attrs: { lang: 'de' } }])));
    expect(h).toContain('<a ');
    expect(h).toContain('lang="de"');
    expect(h).toContain('top:-0.462em');
  });

  it('ignores crafted values instead of passing them to CSS', () => {
    const h = render(doc(run('x', [{ type: 'caps', attrs: { kind: 'all;background:url(//evil)' } }, { type: 'charSpacing', attrs: { px: 'x;color:red' } }, { type: 'lang', attrs: { lang: '"><b>' } }])));
    expect(h).not.toContain('evil');
    expect(h).not.toContain('letter-spacing');
    expect(h).not.toContain('lang=');
  });
});

describe('StaticRichText fields', () => {
  const now = new Date(2026, 0, 15, 9, 5);
  const ctx = (slideNumber: number | null): FieldContext => ({ slideNumber, now, locale: 'en-US' });
  const field = (attrs: object, text: string, marks: Array<{ type: string; attrs?: Record<string, unknown> }> = []): RichTextInline =>
    ({ type: 'field', attrs, content: [{ type: 'text', text, marks }] }) as RichTextInline;

  it('shows the slide number for the slide being drawn, with the cached text carrying the formatting', () => {
    const h = render(doc(field({ kind: 'slidenum' }, '‹#›', [{ type: 'bold' }, { type: 'baseline', attrs: { pct: 30 } }])), ctx(3));
    expect(h).toContain('data-field="slidenum"');
    expect(h).toContain('font-weight:700');
    expect(h).toContain('>3<');
  });

  it('shows the cached text without a context, the date for a date field, and footer text', () => {
    expect(render(doc(field({ kind: 'slidenum' }, '5')))).toContain('>5<');
    expect(render(doc(field({ kind: 'slidenum' }, '')))).toContain('>‹#›<');
    expect(render(doc(field({ kind: 'datetime', format: 'date' }, 'old')), ctx(1))).toContain('>Jan 15, 2026<');
    expect(render(doc(field({ kind: 'footer' }, 'Acme')), ctx(1))).toContain('>Acme<');
    expect(render(doc({ type: 'field', attrs: { kind: 'footer' } } as RichTextInline), ctx(1))).toContain('data-field="footer"');
  });

  it('formatting set on the node itself still applies', () => {
    const f = { type: 'field', attrs: { kind: 'footer' }, content: [{ type: 'text', text: 'x' }], marks: [{ type: 'italic' }] } as RichTextInline;
    expect(render(doc(f))).toContain('font-style:italic');
  });
});

describe('SlideView field context', () => {
  beforeAll(() => {
    jest.spyOn(React, 'useSyncExternalStore').mockImplementation((_sub, getSnapshot) => getSnapshot());
  });
  afterAll(() => jest.restoreAllMocks());
  it('numbers fields by the slide position in the deck', () => {
    const deck = new DeckImpl('d', 'Deck');
    const first = deck.getActiveSlideId();
    const second = deck.addSlide({ afterSlideId: first }).id;
    const body = doc({ type: 'field', attrs: { kind: 'slidenum' }, content: [{ type: 'text', text: '‹#›' }] } as RichTextInline);
    deck.addElement(second, { type: 'text', x: 0, y: 0, w: 200, h: 80, richText: body });
    const h = renderToStaticMarkup(
      <DeckProvider deck={deck}>
        <SlideView slideId={second} />
      </DeckProvider>,
    );
    expect(h).toContain('>2<');
    // A slide that is not in the deck has no number: the cached text shows.
    const gone = renderToStaticMarkup(
      <DeckProvider deck={deck}>
        <SlideView slideId="missing" />
      </DeckProvider>,
    );
    expect(gone).not.toContain('data-field');
  });
});

type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const find = (node: React.ReactNode, pred: (e: React.ReactElement<P>) => boolean, out: React.ReactElement<P>[] = []) => {
  React.Children.forEach(node, (c) => {
    if (!React.isValidElement<P>(c)) return;
    if (pred(c)) out.push(c);
    find(c.props.children, pred, out);
  });
  return out;
};
const byId = (id: string) => (e: React.ReactElement<P>) => e.props['data-testid'] === id;

describe('RunFields', () => {
  const none: ResolvedRun = { baselinePct: null, caps: null, spacingPx: null, lang: null };
  let slots: unknown[] = [];
  let n = 0;
  beforeEach(() => {
    slots = [];
    jest.spyOn(React, 'useState').mockImplementation(((init: unknown) => {
      const i = n++;
      if (!(i in slots)) slots[i] = init;
      return [slots[i], (v: unknown) => { slots[i] = v; }];
    }) as never);
  });
  afterEach(() => jest.restoreAllMocks());
  const tree = (run = none, onMark = jest.fn(), onInsertField?: (f: FieldInsert) => void) => {
    n = 0;
    return { onMark, el: RunFields({ run, onMark, onInsertField }) as React.ReactElement<P> };
  };
  const one = (el: React.ReactNode, id: string) => find(el, byId(id))[0];
  const change = (el: React.ReactNode, id: string, value: string) => one(el, id).props.onChange({ target: { value } });

  it('is a labelled group', () => {
    expect(tree().el.props['aria-label']).toBe('Run options');
  });

  it('baseline: super, sub, normal, and a custom raise is shown without being applied', () => {
    const t = tree();
    change(t.el, 'run-baseline', 'super');
    change(t.el, 'run-baseline', 'sub');
    change(t.el, 'run-baseline', 'normal');
    change(t.el, 'run-baseline', 'custom');
    expect(t.onMark.mock.calls).toEqual([
      [{ type: 'baseline', attrs: { pct: SUPERSCRIPT_PCT } }],
      [{ type: 'baseline', attrs: { pct: SUBSCRIPT_PCT } }],
      [{ type: 'baseline', attrs: null }],
    ]);
    const values = [null, SUPERSCRIPT_PCT, SUBSCRIPT_PCT, 40].map((pct) => one(tree({ ...none, baselinePct: pct }).el, 'run-baseline').props.value);
    expect(values).toEqual(['normal', 'super', 'sub', 'custom']);
    const options = one(tree({ ...none, baselinePct: 40 }).el, 'run-baseline').props.children as Array<React.ReactElement<P> | false>;
    expect((options.find((o) => o && o.props.value === 'custom') as React.ReactElement<P>).props.children).toBe('Raised 40%');
  });

  it('caps and spacing', () => {
    const t = tree({ ...none, caps: 'small', spacingPx: 2 });
    expect(one(t.el, 'run-caps').props.value).toBe('small');
    expect(one(t.el, 'run-spacing').props.value).toBe(2);
    change(t.el, 'run-caps', 'all');
    change(t.el, 'run-caps', 'none');
    one(t.el, 'run-spacing').props.onValue(1.5);
    one(t.el, 'run-spacing').props.onValue(0);
    expect(t.onMark.mock.calls).toEqual([
      [{ type: 'caps', attrs: { kind: 'all' } }],
      [{ type: 'caps', attrs: null }],
      [{ type: 'charSpacing', attrs: { px: 1.5 } }],
      [{ type: 'charSpacing', attrs: null }],
    ]);
    expect(one(tree().el, 'run-caps').props.value).toBe('none');
  });

  it('language commits on blur or Enter, only when well formed, and clears when emptied', () => {
    const t = tree({ ...none, lang: 'de' });
    const lang = () => one(t.el, 'run-lang');
    expect(lang().props.value).toBe('de');
    expect(lang().props['aria-invalid']).toBe(false);
    lang().props.onBlur(); // nothing typed: no change
    lang().props.onKeyDown({ key: 'Enter' });
    expect(t.onMark).not.toHaveBeenCalled();
    lang().props.onChange({ target: { value: 'fr-CA' } });
    n = 0;
    one(RunFields({ run: { ...none, lang: 'de' }, onMark: t.onMark }), 'run-lang').props.onKeyDown({ key: 'Enter' });
    expect(t.onMark).toHaveBeenLastCalledWith({ type: 'lang', attrs: { lang: 'fr-CA' } });
    // a malformed tag is flagged and not applied
    n = 0;
    one(RunFields({ run: none, onMark: t.onMark }), 'run-lang').props.onChange({ target: { value: 'x_y' } });
    n = 0;
    const bad = RunFields({ run: none, onMark: t.onMark });
    expect(one(bad, 'run-lang').props['aria-invalid']).toBe(true);
    one(bad, 'run-lang').props.onBlur();
    expect(t.onMark).toHaveBeenCalledTimes(1);
    // emptied: removes the mark
    n = 0;
    one(RunFields({ run: { ...none, lang: 'de' }, onMark: t.onMark }), 'run-lang').props.onChange({ target: { value: '  ' } });
    n = 0;
    one(RunFields({ run: { ...none, lang: 'de' }, onMark: t.onMark }), 'run-lang').props.onKeyDown({ key: 'Enter' });
    expect(t.onMark).toHaveBeenLastCalledWith({ type: 'lang', attrs: null });
    one(RunFields({ run: none, onMark: t.onMark }), 'run-lang').props.onKeyDown({ key: 'a' });
    expect(t.onMark).toHaveBeenCalledTimes(2);
  });

  it('insert field is disabled when not editing and inserts the chosen field', () => {
    expect(one(tree().el, 'run-insert-field').props.disabled).toBe(true);
    const t = tree(none, jest.fn(), jest.fn());
    const sel = one(t.el, 'run-insert-field');
    expect(sel.props.disabled).toBe(false);
    sel.props.onChange({ target: { value: 'nope' } });
    const onInsert = jest.fn();
    n = 0;
    const wired = RunFields({ run: none, onMark: jest.fn(), onInsertField: onInsert });
    one(wired, 'run-insert-field').props.onChange({ target: { value: 'datetime' } });
    expect(onInsert).toHaveBeenCalledWith({ kind: 'datetime', format: 'datetime' });
    // not editing: choosing does nothing
    one(tree().el, 'run-insert-field').props.onChange({ target: { value: 'slidenum' } });
    expect(FIELD_CHOICES.map((c) => c.v)).toEqual(['slidenum', 'date', 'time', 'datetime']);
  });
});

describe('insertField', () => {
  const makeView = (marks: Array<{ type: string; attrs?: Record<string, unknown> }> = []) => {
    const d = slidesSchema.nodeFromJSON(doc(run('ab', marks)));
    // A stand-in for the live view (no DOM here): applies each transaction to its state.
    const view = {
      state: EditorState.create({ schema: slidesSchema, doc: d }),
      dispatch(tr: Parameters<EditorState['apply']>[0]) {
        view.state = view.state.apply(tr);
      },
      focus: jest.fn(),
    };
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 2)));
    return view as unknown as EditorView & { state: EditorState };
  };
  it('inserts a slide-number field at the caret, carrying the caret marks on its text', () => {
    const view = makeView([{ type: 'bold' }]);
    insertField(view, { kind: 'slidenum' });
    const json = view.state.doc.toJSON().content[0].content;
    const f = json.find((n: { type: string }) => n.type === 'field');
    expect(f).toMatchObject({ attrs: { kind: 'slidenum', format: null }, content: [{ text: '‹#›', marks: [{ type: 'bold' }] }] });
  });
  it('a date field starts with the date; a stored mark takes precedence over the caret marks', () => {
    const view = makeView();
    view.dispatch(view.state.tr.addStoredMark(slidesSchema.marks.italic.create()));
    insertField(view, { kind: 'datetime', format: 'date' });
    const f = view.state.doc.toJSON().content[0].content.find((n: { type: string }) => n.type === 'field');
    expect(f.attrs.format).toBe('date');
    expect(f.content[0].text).toMatch(/\d{4}/);
    expect(f.content[0].marks).toEqual([{ type: 'italic' }]);
  });
  it('a footer field without text has no content', () => {
    const view = makeView();
    insertField(view, { kind: 'footer' });
    const f = view.state.doc.toJSON().content[0].content.find((n: { type: string }) => n.type === 'field');
    expect(f.content).toBeUndefined();
  });
});
