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
 * `alpha` are fractions (`lumMod val="75000"` is 0.75); `alpha` is left out when the colour is opaque.
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

/** A rectangle in px (1 px = 9525 EMU); `rot` is degrees clockwise and left out when 0. */
export interface PptxFrame {
  x: number;
  y: number;
  w: number;
  h: number;
  rot?: number;
}

/** A bullet as the file writes it: none, a literal glyph or an auto-number scheme (`buAutoNum type`, e.g. `arabicPeriod`). */
export type PptxBullet = { kind: 'none' } | { kind: 'char'; char: string } | { kind: 'auto'; scheme: string; startAt?: number };

/**
 * The text defaults of one paragraph level (`a:lvl1pPr` ... `a:lvl9pPr` with its `a:defRPr`), already merged down the
 * inheritance chain, so a level holds what applies to a paragraph of that level. Font size is in pt, `marginLeft` and
 * `indent` in px, `lineHeight` a multiplier (`spcPct`), `spaceBefore` / `spaceAfter` in pt. A property the chain never sets is absent.
 */
export interface PptxLevelStyle {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  color?: PptxColor;
  /** `'major'` / `'minor'` for the theme fonts (`+mj-lt`, `+mn-lt`), else a family name. */
  font?: string;
  align?: 'left' | 'center' | 'right' | 'justify';
  marginLeft?: number;
  indent?: number;
  lineHeight?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  bullet?: PptxBullet;
}

/** The nine paragraph levels of a placeholder; a level the chain never mentions is `{}`. */
export type PptxLevelStyles = PptxLevelStyle[];

/**
 * A placeholder of a layout, flattened (plan K4, decision SD6): `frame` is the layout's own position, else the one it
 * inherits from the matching master placeholder; `levels` holds the effective defaults of its text.
 */
export interface PptxPlaceholder {
  /** `p:ph/@type` as written (`title`, `ctrTitle`, `subTitle`, `body`, `dt`, `ftr`, `sldNum`, `pic`, ...), `obj` when absent. */
  type: string;
  /** `p:ph/@idx`; absent for the master-style placeholders that carry none. */
  idx?: number;
  name?: string;
  frame?: PptxFrame;
  /** Vertical anchor of the text body. */
  anchor?: 'top' | 'middle' | 'bottom';
  levels: PptxLevelStyles;
  /** The prompt text of the layout placeholder. */
  prompt?: string;
}

/**
 * A slide layout flattened with its master. `elements` of the editor's `LayoutData` (static shapes) arrive with the slide
 * shape reader; here `staticShapes` only counts the non-placeholder shapes the layout and its master draw.
 */
export interface PptxLayout {
  partName: string;
  name: string;
  /** `type` of `p:sldLayout` (`title`, `obj`, `blank`...), when set. */
  layoutType?: string;
  masterPart: string;
  /** The master's own name (`p:cSld/@name`), when it has one. */
  masterName?: string;
  /** The layout's background (else the master's) when it is a solid colour. */
  background?: PptxColor;
  /** `showMasterSp="0"`: the layout hides the master's shapes. */
  hideMasterShapes?: boolean;
  placeholders: PptxPlaceholder[];
  staticShapes: number;
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
  /** Every layout of every master, in master order then layout order, flattened. */
  layouts: PptxLayout[];
  sections?: PptxSection[];
  warnings: string[];
}
