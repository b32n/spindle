// PPTX interchange shapes shared by `@b32nio/spindle-io/pptx` and the slides packages. Like `export/docx` they are
// deliberately independent of slides-core's scene model (shared sits below the editor packages and cannot import it);
// `slides-core` maps a `PresentationImportResult` to a deck (`importFromPptx`, plan K15). The reader (plan K9) grows
// this result in steps: package, presentation and theme first, then layouts, then slide content.

/** Options for the loss detector (`detectPptxLoss`). All optional. */
export interface DetectPptxLossOptions {
  /** `'summary'` (default) drives the import-confirmation dialog; `'full'` adds the finer ids the fidelity harness records. */
  detail?: 'summary' | 'full';
}

/** The 12 theme colour slots of a DrawingML theme, named as in the file. */
export type PptxThemeSlot =
  | 'dk1'
  | 'lt1'
  | 'dk2'
  | 'lt2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hlink'
  | 'folHlink';

/**
 * A colour as the file writes it, with the transforms the editor can carry: a theme slot (after the colour map has
 * turned `bg1`, `tx1` and friends into a slot) with `lumMod` / `lumOff` / `alpha`, or a literal. `lumMod`, `lumOff` and
 * `alpha` are fractions (`lumMod val="75000"` is 0.75).
 */
export type PptxColor =
  | { kind: 'theme'; slot: PptxThemeSlot; alpha?: number; lumMod?: number; lumOff?: number }
  | { kind: 'rgb'; hex: string; alpha?: number };

/** A theme: 12 colours as `#RRGGBB` and the major (headings) and minor (body) font families. */
export interface PptxTheme {
  name: string;
  colors: Record<PptxThemeSlot, string>;
  fonts: {
    major: string;
    minor: string;
    /** East Asian and complex-script families of the font scheme, when it names one. */
    majorEastAsian?: string;
    minorEastAsian?: string;
    majorComplex?: string;
    minorComplex?: string;
  };
}

/** The role-to-slot colour map of a master (`p:clrMap`): `bg1` -> `lt1`, `tx1` -> `dk1` by default. */
export type PptxColorMap = Record<'bg1' | 'tx1' | 'bg2' | 'tx2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6' | 'hlink' | 'folHlink', PptxThemeSlot>;

/** A named group of slides (`p14:section`). `slideIds` are the `pptxId`s of its slides, in order. */
export interface PptxSection {
  name: string;
  slideIds: string[];
}

/**
 * One slide of the presentation, in presentation order. `pptxId` is the slide's `p:sldId` (kept so timing and section
 * references can be re-emitted); `partName` is where it was read from. A slide whose part is missing or damaged is
 * kept as a labelled placeholder: `error` says why and the importer adds a warning.
 */
export interface PptxSlide {
  pptxId: string;
  partName: string;
  /** The slide's own name (`p:cSld/@name`), when it has one. */
  name?: string;
  /** `show="0"`: skipped when presenting. */
  hidden?: boolean;
  /** The layout part the slide uses (`ppt/slideLayouts/slideLayout2.xml`). */
  layoutPart?: string;
  error?: string;
}

/** What `importPptx` returns. Warnings are plain sentences; nothing in them is a failure (a failure is an `ImportError`). */
export interface PresentationImportResult {
  /** Slide size in px (1 px = 9525 EMU, so 12192000 EMU = 1280 px). */
  slideSize: { w: number; h: number };
  /** The theme of the first master (a deck has one theme in the editor). */
  theme: PptxTheme;
  /** The colour map of the first master. */
  colorMap: PptxColorMap;
  slides: PptxSlide[];
  sections?: PptxSection[];
  warnings: string[];
}
