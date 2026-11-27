# Data Structures

A deck is plain JSON — easy to author by hand, generate with an LLM, or diff. `deck.getData()` emits a fully-populated `DeckData`; `deck.setData(data)` accepts it (and normalizes hand-authored input, filling in missing ids and fractional indices).

## DeckData

```ts
interface DeckData {
  id: string;
  title: string;
  slideSize: { w: number; h: number };   // default 1280×720 (16:9 @ 96 DPI)
  theme: ThemeData;                        // 12 color slots + major/minor fonts
  layouts?: LayoutData[];                  // defaults to the built-ins
  slides: Array<Slide & { elements: SlideElement[] }>;
  threads?: SlidesCommentThread[];
  selection?: { slideId?: string; elementIds: string[] };  // round-trips; not CRDT state
}
```

## Slides & elements

```ts
interface Slide {
  id: string;
  index: string;          // fractional order key
  layoutRef?: string;
  background?: Fill;
  notes?: RichTextDoc;
}

interface ElementBase {
  id: string;
  containerId: string;    // the slide (or, later, board) this belongs to
  index: string;          // fractional z-order key (ascending = back-to-front)
  x: number; y: number; w: number; h: number;
  rotation: number;       // degrees, clockwise, about the box centre
  groupId?: string; parentGroupId?: string;   // innermost group; the group that group is nested in
  locked?: boolean; opacity?: number;
  placeholder?: { type: PlaceholderType; idx: number };   // PPTX fidelity
}
```

The element union is `text` · `shape` · `image` · `line`:

- **`TextElement`** — `richText` (ProseMirror JSON), `bodyStyle`, optional `fill` / `stroke`.
- **`ShapeElement`** — `shape` (one of 76 presets, `unknown:<prst>` for an OOXML preset the editor cannot draw, kept for re-export, or `path` for custom geometry held in `path`, a `CustomGeometry` of paths of move/line/cubic/quad/arc/close commands), `fill`, optional `stroke`, `adjustments`, `flipH/V`, and its own optional `richText`.
- **`ImageElement`** — `src`, `naturalW/H`, `flipH/V`.
- **`LineElement`** — `stroke`, `startArrow?` / `endArrow?`, `flipH/V`, `startBind?` / `endBind?` (anchor on another element), and for connectors `connector?` (`'elbow'` or `'curved'`; absent = straight) with `adjustments?` (elbow middle-leg position, 0 to 1). Shape-specific connection sites (ellipse, diamond, triangle) reuse the eight anchor ids; see `scene/connector.ts`.

## Colors, fills, strokes

Colors are symbolic by default, so a theme switch recolors the deck at render time:

```ts
type Color =
  | { kind: 'theme'; slot: ThemeColorSlot; alpha?; lumMod?; lumOff? }
  | { kind: 'rgb'; hex: string; alpha? };

type Fill = { kind: 'none' } | { kind: 'solid'; color: Color };
interface Stroke { color: Color; width: number; dash?: 'solid' | 'dash' | 'dot' }
```

The 12 `ThemeColorSlot`s are PPTX-compatible: `dk1 lt1 dk2 lt2 accent1…6 hlink folHlink`.

## Rich text

Text bodies are a flat ProseMirror document — `doc > paragraph+ > inline*` — matching the way PPTX stores text (flat paragraphs with bullet attributes, not nested lists). Paragraph attrs carry `align`, `listType`, `indent`, `lineHeight`, and spacing; marks are `bold`, `italic`, `underline`, `strikethrough`, `link`, `textColor` (a `Color`, so theme slots survive), `fontFamily`, and `fontSize`. Element-level defaults on `bodyStyle` (the PPTX `defRPr` analogue) let placeholder prompts and freshly-typed text inherit the intended size/color.

## Fractional indices

Both slide order and element z-order use fractional-index strings, never array position — Yjs has no move operation, so reordering rewrites one key rather than shuffling an array. Keys are base-62 lexicographic with a random jitter suffix; ties break by id.
