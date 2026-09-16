import { useMemo } from 'react';
import { DeckProvider, SlidesEditor } from '@b32nio/spindle-slides-react';
import { DeckImpl, richTextFromPlainText, getBuiltinTheme } from '@b32nio/spindle-slides-core';
import type { DeckData } from '@b32nio/spindle-slides-core';
import DemoChrome from './DemoChrome';
import { exportDeckToPdf } from './export-pdf';

// A small three-slide deck rendered by the real Spindle Slides engine: text,
// shapes, an inline-SVG image, a line, and themed colors. Select an element and
// drag/resize/rotate it; double-click text to edit; press Present or export a PDF.

const base = { id: '', containerId: '', index: '' };
const accent = (n: 1 | 2 | 3 | 4 | 5 | 6) => ({ kind: 'theme' as const, slot: `accent${n}` as const });

const IMG =
  'data:image/svg+xml;utf8,' +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="480" height="300"><defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#2D7FF9"/><stop offset="1" stop-color="#16B1A6"/></linearGradient></defs><rect width="480" height="300" fill="url(#g)"/><circle cx="370" cy="80" r="46" fill="#fff" opacity="0.85"/><rect x="36" y="200" width="280" height="16" rx="8" fill="#fff" opacity="0.8"/><rect x="36" y="230" width="180" height="16" rx="8" fill="#fff" opacity="0.6"/></svg>`
  );

function makeDeck(): DeckImpl {
  const data: DeckData = {
    id: 'spindle-slides-demo',
    title: 'Spindle Slides — Demo',
    slideSize: { w: 1280, h: 720 },
    theme: getBuiltinTheme('Clean'),
    slides: [
      {
        id: '', index: '', layoutRef: 'title',
        elements: [
          { ...base, type: 'shape', shape: 'rect', x: 0, y: 560, w: 1280, h: 160, rotation: 0, fill: { kind: 'solid', color: accent(1) }, opacity: 0.12 },
          { ...base, type: 'text', x: 160, y: 250, w: 960, h: 150, rotation: 0, richText: { type: 'doc', content: [{ type: 'paragraph', attrs: { align: 'center' }, content: [{ type: 'text', text: 'Spindle ', marks: [{ type: 'bold' }, { type: 'fontSize', attrs: { size: 66 } }] }, { type: 'text', text: 'Slides', marks: [{ type: 'bold' }, { type: 'fontSize', attrs: { size: 66 } }, { type: 'textColor', attrs: { color: accent(1) } }] }] }] }, bodyStyle: { vAlign: 'bottom' } },
          { ...base, type: 'text', x: 160, y: 410, w: 960, h: 70, rotation: 0, richText: { type: 'doc', content: [{ type: 'paragraph', attrs: { align: 'center' }, content: [{ type: 'text', text: 'An AI-first, collaborative presentation engine', marks: [{ type: 'fontSize', attrs: { size: 26 } }, { type: 'textColor', attrs: { color: { kind: 'theme', slot: 'dk2' } } }] }] }] } },
        ],
      },
      {
        id: '', index: '', layoutRef: 'titleContent',
        elements: [
          { ...base, type: 'text', x: 80, y: 48, w: 760, h: 90, rotation: 0, richText: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: "What's in the box", marks: [{ type: 'bold' }, { type: 'fontSize', attrs: { size: 40 } }] }] }] }, bodyStyle: { vAlign: 'middle' } },
          { ...base, type: 'text', x: 80, y: 180, w: 640, h: 460, rotation: 0, richText: { type: 'doc', content: ['Text, shapes, images and lines', 'Move, resize, rotate — with smart guides', '12-slot themes, present mode, PDF export', 'Real-time collaboration over Yjs', 'Element-anchored comments'].map((t) => ({ type: 'paragraph', attrs: { listType: 'bullet', spaceAfter: 8 }, content: [{ type: 'text', text: t, marks: [{ type: 'fontSize', attrs: { size: 24 } }] }] })) } },
          { ...base, type: 'shape', shape: 'ellipse', x: 800, y: 190, w: 180, h: 180, rotation: 0, fill: { kind: 'solid', color: accent(2) } },
          { ...base, type: 'shape', shape: 'roundRect', x: 1000, y: 190, w: 180, h: 180, rotation: 12, fill: { kind: 'solid', color: accent(3) } },
          { ...base, type: 'shape', shape: 'star5', x: 800, y: 420, w: 180, h: 180, rotation: 0, fill: { kind: 'solid', color: accent(4) } },
          { ...base, type: 'shape', shape: 'hexagon', x: 1000, y: 420, w: 180, h: 180, rotation: 0, fill: { kind: 'solid', color: accent(5) } },
        ],
      },
      {
        id: '', index: '', layoutRef: 'blank',
        background: { kind: 'solid', color: { kind: 'theme', slot: 'lt2' } },
        elements: [
          { ...base, type: 'text', x: 80, y: 40, w: 1120, h: 70, rotation: 0, richText: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'Shapes, images & lines', marks: [{ type: 'bold' }, { type: 'fontSize', attrs: { size: 34 } }] }] }] }, bodyStyle: { vAlign: 'middle' } },
          { ...base, type: 'shape', shape: 'triangle', x: 80, y: 150, w: 150, h: 120, rotation: 0, fill: { kind: 'solid', color: accent(1) } },
          { ...base, type: 'shape', shape: 'diamond', x: 260, y: 150, w: 150, h: 120, rotation: 0, fill: { kind: 'solid', color: accent(2) } },
          { ...base, type: 'shape', shape: 'arrowRight', x: 440, y: 150, w: 150, h: 120, rotation: 0, fill: { kind: 'solid', color: accent(3) } },
          { ...base, type: 'shape', shape: 'cloud', x: 620, y: 150, w: 150, h: 120, rotation: 0, fill: { kind: 'solid', color: accent(5) } },
          { ...base, type: 'shape', shape: 'heart', x: 800, y: 150, w: 150, h: 120, rotation: 0, fill: { kind: 'solid', color: accent(4) } },
          { ...base, type: 'shape', shape: 'roundRect', x: 80, y: 330, w: 340, h: 120, rotation: 0, fill: { kind: 'solid', color: { kind: 'theme', slot: 'accent1', lumMod: 1, lumOff: 0.3 } }, stroke: { color: accent(1), width: 2 }, richText: richTextFromPlainText('Shapes hold text too'), bodyStyle: { vAlign: 'middle' } },
          { ...base, type: 'image', x: 470, y: 320, w: 360, h: 225, rotation: 0, src: IMG, naturalW: 480, naturalH: 300 },
          { ...base, type: 'line', x: 890, y: 340, w: 300, h: 200, rotation: 0, stroke: { color: accent(6), width: 4 }, endArrow: 'triangle' },
        ],
      },
    ],
  };
  const deck = new DeckImpl('spindle-slides-demo', 'Spindle Slides — Demo');
  deck.setData(data);
  return deck;
}

const DEMO_USERS = [
  { id: 'alice', name: 'Alice' },
  { id: 'bob', name: 'Bob' },
];

export default function SlidesDemo() {
  const deck = useMemo(makeDeck, []);
  return (
    <DemoChrome active="slides" hint="Real engine · drag, edit, present, export">
      {({ width, height }) => (
        <div style={{ width, height, overflow: 'hidden', background: '#fff' }}>
          <DeckProvider deck={deck} currentUser={DEMO_USERS[0]} mentionableUsers={DEMO_USERS}>
            <SlidesEditor
              headerActions={
                <button
                  onClick={() => exportDeckToPdf(deck)}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: 6, border: '1px solid #d5d9e0', background: '#fff', color: '#3e4c59', borderRadius: 5, padding: '6px 12px', fontSize: 13, cursor: 'pointer' }}
                >
                  PDF
                </button>
              }
            />
          </DeckProvider>
        </div>
      )}
    </DemoChrome>
  );
}
