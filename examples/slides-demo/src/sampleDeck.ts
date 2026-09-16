// A hand-authored 3-slide deck exercising every element type, several shape
// presets, themed colors, and rich-text marks. Ids and fractional indices are
// left blank — normalizeDeckData (called by setData) fills them in.

import {
  getBuiltinTheme,
  DEFAULT_SLIDE_SIZE,
  richTextFromPlainText,
  type DeckData,
  type SlideElement,
  type RichTextDoc,
  type RichTextInline,
  type Color,
} from '@b32nio/spindle-slides-core';

// Base fields every element needs; normalization replaces the empty strings.
const base = { id: '', containerId: '', index: '' };

type Frame = { x: number; y: number; w: number; h: number; rotation?: number };

function run(text: string, marks?: RichTextInline['marks']): RichTextInline {
  return marks ? { type: 'text', text, marks } : { type: 'text', text };
}

function doc(paragraphs: RichTextDoc['content']): RichTextDoc {
  return { type: 'doc', content: paragraphs };
}

function bullets(items: string[]): RichTextDoc {
  return doc(
    items.map((text) => ({
      type: 'paragraph',
      attrs: { listType: 'bullet', spaceAfter: 8 },
      content: [run(text)],
    }))
  );
}

function text(frame: Frame, rich: RichTextDoc, extra: Partial<SlideElement> = {}): SlideElement {
  return { ...base, type: 'text', rotation: 0, ...frame, richText: rich, ...extra } as SlideElement;
}

function shape(
  frame: Frame,
  preset: string,
  color: Color,
  extra: Partial<SlideElement> = {}
): SlideElement {
  return {
    ...base,
    type: 'shape',
    rotation: 0,
    ...frame,
    shape: preset,
    fill: { kind: 'solid', color },
    ...extra,
  } as SlideElement;
}

const accent = (n: 1 | 2 | 3 | 4 | 5 | 6): Color => ({ kind: 'theme', slot: `accent${n}` });

// A self-contained inline-SVG "image" so the demo works offline.
const SAMPLE_IMAGE =
  'data:image/svg+xml;utf8,' +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="480" height="320">
      <defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0" stop-color="#2D7FF9"/><stop offset="1" stop-color="#16B1A6"/>
      </linearGradient></defs>
      <rect width="480" height="320" fill="url(#g)"/>
      <circle cx="360" cy="90" r="54" fill="#ffffff" opacity="0.85"/>
      <rect x="40" y="210" width="300" height="18" rx="9" fill="#ffffff" opacity="0.8"/>
      <rect x="40" y="244" width="200" height="18" rx="9" fill="#ffffff" opacity="0.6"/>
    </svg>`
  );

export function buildSampleDeck(): DeckData {
  return {
    id: 'demo-deck',
    title: 'Spindle Slides — Demo',
    slideSize: { ...DEFAULT_SLIDE_SIZE },
    theme: getBuiltinTheme('Clean'),
    slides: [
      // ── Slide 1: title ──────────────────────────────────────────────────
      {
        id: '',
        index: '',
        layoutRef: 'title',
        elements: [
          shape({ x: 0, y: 560, w: 1280, h: 160 }, 'rect', accent(1), { opacity: 0.12 }),
          text(
            { x: 160, y: 250, w: 960, h: 150 },
            doc([
              {
                type: 'paragraph',
                attrs: { align: 'center' },
                content: [
                  run('Spindle ', [
                    { type: 'bold' },
                    { type: 'fontSize', attrs: { size: 66 } },
                  ]),
                  run('Slides', [
                    { type: 'bold' },
                    { type: 'fontSize', attrs: { size: 66 } },
                    { type: 'textColor', attrs: { color: accent(1) } },
                  ]),
                ],
              },
            ]),
            { bodyStyle: { vAlign: 'bottom', padding: 8 } }
          ),
          text(
            { x: 160, y: 410, w: 960, h: 70 },
            doc([
              {
                type: 'paragraph',
                attrs: { align: 'center' },
                content: [
                  run('An AI-first, collaborative presentation engine', [
                    { type: 'fontSize', attrs: { size: 26 } },
                    { type: 'textColor', attrs: { color: { kind: 'theme', slot: 'dk2' } } },
                  ]),
                ],
              },
            ]),
            { bodyStyle: { vAlign: 'top', padding: 8 } }
          ),
        ],
      },
      // ── Slide 2: title + content + accents ──────────────────────────────
      {
        id: '',
        index: '',
        layoutRef: 'titleContent',
        elements: [
          text(
            { x: 80, y: 48, w: 800, h: 90 },
            doc([
              {
                type: 'paragraph',
                content: [run("What's in the box", [{ type: 'bold' }, { type: 'fontSize', attrs: { size: 40 } }])],
              },
            ]),
            { bodyStyle: { vAlign: 'middle', padding: 8 } }
          ),
          text(
            { x: 80, y: 180, w: 640, h: 460 },
            (() => {
              const b = bullets([
                'Text, shapes, images and lines',
                'Multi-select, move, resize and rotate',
                'Themes with 12 symbolic color slots',
                'Real-time collaboration over Yjs',
                'Present mode and PDF export',
              ]);
              b.content.forEach((p) => (p.content = p.content?.map((r) => run(r.text, [{ type: 'fontSize', attrs: { size: 24 } }]))));
              return b;
            })(),
            { bodyStyle: { vAlign: 'top', padding: 8 } }
          ),
          shape({ x: 800, y: 190, w: 180, h: 180 }, 'ellipse', accent(2)),
          shape({ x: 1000, y: 190, w: 180, h: 180, rotation: 12 }, 'roundRect', accent(3)),
          shape({ x: 800, y: 420, w: 180, h: 180 }, 'star5', accent(4)),
          shape({ x: 1000, y: 420, w: 180, h: 180 }, 'hexagon', accent(5)),
        ],
      },
      // ── Slide 3: shapes gallery + image + line ──────────────────────────
      {
        id: '',
        index: '',
        layoutRef: 'blank',
        background: { kind: 'solid', color: { kind: 'theme', slot: 'lt2' } },
        elements: [
          text(
            { x: 80, y: 40, w: 1120, h: 70 },
            doc([{ type: 'paragraph', content: [run('Shapes, images & lines', [{ type: 'bold' }, { type: 'fontSize', attrs: { size: 34 } }])] }]),
            { bodyStyle: { vAlign: 'middle', padding: 8 } }
          ),
          shape({ x: 80, y: 150, w: 150, h: 120 }, 'triangle', accent(1)),
          shape({ x: 260, y: 150, w: 150, h: 120 }, 'diamond', accent(2)),
          shape({ x: 440, y: 150, w: 150, h: 120 }, 'arrowRight', accent(3)),
          shape({ x: 620, y: 150, w: 150, h: 120 }, 'chevron', accent(4)),
          shape({ x: 800, y: 150, w: 150, h: 120 }, 'cloud', accent(5)),
          shape({ x: 980, y: 150, w: 150, h: 120 }, 'heart', accent(4)),
          shape(
            { x: 80, y: 330, w: 320, h: 120 },
            'roundRect',
            { kind: 'theme', slot: 'accent1', lumMod: 1, lumOff: 0.3 },
            {
              richText: richTextFromPlainText('Shapes hold text too'),
              bodyStyle: { vAlign: 'middle', padding: 8 },
              stroke: { color: accent(1), width: 2 },
            }
          ),
          {
            ...base,
            type: 'image',
            rotation: 0,
            x: 460,
            y: 320,
            w: 360,
            h: 240,
            src: SAMPLE_IMAGE,
            naturalW: 480,
            naturalH: 320,
          } as SlideElement,
          {
            ...base,
            type: 'line',
            rotation: 0,
            x: 880,
            y: 340,
            w: 300,
            h: 200,
            stroke: { color: accent(6), width: 4 },
            endArrow: 'triangle',
          } as SlideElement,
        ],
      },
    ],
  };
}
