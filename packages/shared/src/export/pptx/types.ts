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
  /** The slide's own background (`p:bg`), when it has one. */
  background?: PptxFill;
  /** The `p:sp` shapes, `p:pic` pictures and table `p:graphicFrame`s in drawing order; the other shape kinds arrive with the later reader steps. Absent for a slide that failed. */
  elements?: PptxElement[];
  error?: string;
}

/** A rectangle in px (1 px = 9525 EMU); `rot` is degrees clockwise and left out when 0; the flips are left out when off. */
export interface PptxFrame {
  x: number;
  y: number;
  w: number;
  h: number;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
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
  /** Exact line spacing in pt (`lnSpc/spcPts`); excludes `lineHeight`. */
  lineSpacingPts?: number;
  spaceBefore?: number;
  /** Space before as a fraction of the font size (`spcBef/spcPct`, `0.2` is 20%); excludes `spaceBefore`. */
  spaceBeforePct?: number;
  spaceAfter?: number;
  /** Space after as a fraction of the font size (`spcAft/spcPct`); excludes `spaceAfter`. */
  spaceAfterPct?: number;
  bullet?: PptxBullet;
  /** `buClr`; absent when the bullet follows the text colour. */
  bulletColor?: PptxColor;
  /** `buSzPct` as a fraction of the text size (`1` is 100%). */
  bulletSizePct?: number;
  /** `buFont` family. */
  bulletFont?: string;
  tabStops?: PptxTabStop[];
}

/** A tab stop (`a:tab`): `pos` in px, `align` as written (`l`, `ctr`, `r`, `dec`). */
export interface PptxTabStop {
  pos: number;
  align: string;
}

/** The nine paragraph levels of a placeholder; a level the chain never mentions is `{}`. */
export type PptxLevelStyles = PptxLevelStyle[];

/**
 * The attributes of a text body's `a:bodyPr` that the editor models (plan K3a), only the ones the element (or, for a
 * placeholder, its chain) states. Insets are in px.
 */
export interface PptxBodyProps {
  insets?: { l?: number; t?: number; r?: number; b?: number };
  anchor?: 'top' | 'middle' | 'bottom';
  /** `anchorCtr`: the text block is centred horizontally as a whole. */
  anchorCtr?: boolean;
  /** `vert` as written (`horz`, `vert`, `vert270`, `eaVert`, `wordArtVert`, `mongolianVert`...). */
  vert?: string;
  /** `wrap="none"` is `none`. */
  wrap?: 'square' | 'none';
  /** The autofit mode with the file's own scale: `fontScale` and `lnSpcReduction` are fractions (`62500` is 0.625). */
  autofit?: { kind: 'none' | 'shrink' | 'resize'; fontScale?: number; lnSpcReduction?: number };
  columns?: number;
  /** Column gap in px. */
  columnSpacing?: number;
  /** Text rotation inside the shape in degrees. */
  rot?: number;
}

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
  /** The body properties the layout placeholder states, over the master placeholder's. */
  bodyProps?: PptxBodyProps;
  levels: PptxLevelStyles;
  /** The prompt text of the layout placeholder. */
  prompt?: string;
}

/** A fill as the file writes it (plan K1): none, a solid colour, a gradient, a pattern or a picture. */
export type PptxFill =
  | { kind: 'none' }
  | { kind: 'solid'; color: PptxColor }
  | {
      kind: 'gradient';
      /** `linear` (`a:lin`), `radial` (`a:path path="circle"`), `rect` or `shape` (the other paths). */
      shape: 'linear' | 'radial' | 'rect' | 'shape';
      /** Degrees clockwise from the x axis, for a linear gradient. */
      angle?: number;
      /** `a:lin scaled="1"`. */
      scaled?: boolean;
      /** `rotWithShape="1"`. */
      rotWithShape?: boolean;
      /** Stops in order; `pos` is a fraction (0 to 1). */
      stops: { pos: number; color: PptxColor }[];
    }
  | { kind: 'pattern'; prst: string; fg?: PptxColor; bg?: PptxColor }
  /** `a:blipFill`: `part` is the image part it points at (the bytes arrive with the picture reader, K10). */
  | { kind: 'picture'; part?: string; tile?: boolean; alpha?: number };

/** A line end decoration (`a:headEnd`, `a:tailEnd`) as written. */
export interface PptxLineEnd {
  type: string;
  w?: string;
  len?: string;
}

/** An outline (`a:ln`): `width` in px; the rest as written. A property the element does not state is absent. */
export interface PptxLine {
  fill?: PptxFill;
  width?: number;
  /** `prstDash` (`dash`, `sysDot`...) or `custom` for a `custDash`. */
  dash?: string;
  cap?: 'flat' | 'round' | 'square';
  join?: 'round' | 'bevel' | 'miter';
  head?: PptxLineEnd;
  tail?: PptxLineEnd;
}

/**
 * A shape's geometry (plan K5): a preset by its OOXML `prst` name with its adjustments (`known` is false for a preset
 * the editor's library lacks, which the editor keeps under that name and draws approximately), or custom geometry (read by K10).
 */
export type PptxGeometry = { kind: 'preset'; prst: string; adjustments?: Record<string, number>; known: boolean } | { kind: 'custom' };

/** A run's attributes (`a:rPr`) beyond what its paragraph's level style gives; only what the run states. */
export interface PptxRun {
  /** `text` (`a:r`), `field` (`a:fld`: `text` is the cached value) or `break` (`a:br`: no text). */
  kind: 'text' | 'field' | 'break';
  text: string;
  /** `a:fld/@type` (`slidenum`, `datetime1`...). */
  fieldType?: string;
  bold?: boolean;
  italic?: boolean;
  /** `u` as written (`sng`, `dbl`...); `none` is not kept. */
  underline?: string;
  strike?: 'single' | 'double';
  /** Size in pt. */
  size?: number;
  color?: PptxColor;
  highlight?: PptxColor;
  /** `'major'` / `'minor'` for theme fonts, else a family name. */
  font?: string;
  fontEastAsian?: string;
  fontComplex?: string;
  /** Raise (positive) or lower as a fraction of the text size (`baseline="30000"` is 0.3). */
  baseline?: number;
  caps?: 'all' | 'small';
  /** Character spacing in pt. */
  spacing?: number;
  lang?: string;
  /** `a:hlinkClick` on the run. */
  link?: PptxLink;
}

/** A paragraph: its effective level style (presentation default, layout, shape list style and its own `a:pPr` merged) and its runs. */
export interface PptxParagraph {
  /** 0 to 8. */
  level: number;
  style: PptxLevelStyle;
  rtl?: boolean;
  runs: PptxRun[];
  /** Size of the paragraph mark (`a:endParaRPr`) in pt, which sizes an empty paragraph. */
  endSize?: number;
}

export interface PptxTextBody {
  props: PptxBodyProps;
  paragraphs: PptxParagraph[];
}

/**
 * A `p:sp` as the slide (or a layout's static content) draws it. `frame` is effective: a placeholder without its own
 * position takes the layout's. Text is flattened the same way (`PptxParagraph.style`).
 */
export interface PptxShape {
  kind: 'shape';
  /** `p:cNvPr/@id`, kept so timing references can be re-emitted. */
  pptxId: string;
  name: string;
  /** Alternative text (`descr`). */
  altText?: string;
  /** `hidden="1"`. */
  hidden?: boolean;
  /** `p:ph`: kept with the layout it follows. */
  placeholder?: { type: string; idx?: number };
  /** `txBox="1"`. */
  textBox?: boolean;
  frame?: PptxFrame;
  geometry?: PptxGeometry;
  fill?: PptxFill;
  line?: PptxLine;
  text?: PptxTextBody;
  /** `a:hlinkClick` on the shape. */
  link?: PptxLink;
}

/**
 * A hyperlink (`a:hlinkClick`). Only `http`, `https`, `mailto` and `tel` addresses become a `url`; an `action`
 * (`ppaction://` targets such as a jump to a slide or a macro, and `file:` targets) is kept inert, as written, and never
 * followed (`pptx.actions`). An address with any other scheme (`javascript:`, `data:`...) is dropped with a warning.
 */
export type PptxLink = { kind: 'url'; href: string; tooltip?: string } | { kind: 'action'; target: string; slidePart?: string; tooltip?: string };

/**
 * A picture (`p:pic`). `dataUrl` holds the embedded image (never fetched from a link); it is absent when `unsupported` says
 * why the editor cannot show it: `linked` (the file lives outside the package), `missing` (the relationship or part is
 * absent or too large) or the format (`emf`, `wmf`, ... have no browser renderer: kept as a placeholder, the part name is
 * in `part`). An SVG with a PNG fallback reads as the PNG.
 */
export interface PptxPicture {
  kind: 'picture';
  pptxId: string;
  name: string;
  altText?: string;
  hidden?: boolean;
  placeholder?: { type: string; idx?: number };
  frame?: PptxFrame;
  geometry?: PptxGeometry;
  line?: PptxLine;
  /** The image part (`ppt/media/image1.png`). */
  part?: string;
  mime?: string;
  dataUrl?: string;
  /** Percent of the source trimmed from each edge (`a:srcRect`), only when something is cropped. */
  crop?: { l: number; t: number; r: number; b: number };
  /** `a:alphaModFix` as a fraction, when under 1. */
  alpha?: number;
  unsupported?: string;
  link?: PptxLink;
}

/** Cell borders by side; `null` is an explicit "no line" (`a:noFill`), an absent side states nothing. */
export type PptxCellBorders = Partial<Record<'l' | 't' | 'r' | 'b', PptxLine | null>>;

/**
 * One cell of a table. `fill`, `borders` and the text's bold, italic and colour are **resolved**: the cell's own `a:tcPr`
 * and run properties over the table style (`ppt/tableStyles.xml`: whole table, banding, first and last row and column).
 * A merged region is its origin cell with `gridSpan` / `rowSpan`; the cells it covers are kept (`hMerge` / `vMerge`) as in the file.
 */
export interface PptxTableCell {
  text: PptxTextBody;
  gridSpan?: number;
  rowSpan?: number;
  hMerge?: boolean;
  vMerge?: boolean;
  fill?: PptxFill;
  borders?: PptxCellBorders;
  /** Cell margins in px (`marL`, `marT`, `marR`, `marB`), only what the cell states. */
  margins?: { l?: number; t?: number; r?: number; b?: number };
  anchor?: 'top' | 'middle' | 'bottom';
  /** `a:tcPr/@vert` as written. */
  vert?: string;
}

/** A table (`p:graphicFrame` holding `a:tbl`): column widths and row heights in px, rows of cells. */
export interface PptxTable {
  kind: 'table';
  pptxId: string;
  name: string;
  altText?: string;
  hidden?: boolean;
  frame?: PptxFrame;
  columns: number[];
  rows: { height: number; cells: PptxTableCell[] }[];
  /** The `a:tblPr` switches that are on. */
  flags?: Partial<Record<'firstRow' | 'firstCol' | 'lastRow' | 'lastCol' | 'bandRow' | 'bandCol', true>>;
  /** `a:tableStyleId` and the style's name, when the file defines it. */
  styleId?: string;
  styleName?: string;
}

/** What a slide holds, in drawing order. */
export type PptxElement = PptxShape | PptxPicture | PptxTable;

/**
 * A slide layout flattened with its master. `elements` is the static content (the master's shapes, unless the layout hides
 * them, then the layout's own) in drawing order, for the shape kind the reader reads so far (`p:sp`); `staticShapes` counts every
 * non-placeholder shape of either, pictures and groups included.
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
  elements: PptxShape[];
  staticShapes: number;
}

/** The fonts a deck names (`typeface`), without theme references, and those outside the short list the editor can show without the file carrying them. */
export interface PptxFonts {
  used: string[];
  unavailable: string[];
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
  fonts: PptxFonts;
  sections?: PptxSection[];
  warnings: string[];
}
