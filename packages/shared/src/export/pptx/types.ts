// PPTX interchange shapes shared by `@b32nio/spindle-io/pptx` and the slides packages. Like `export/docx` they are
// deliberately independent of slides-core's scene model (shared sits below the editor packages and cannot import it);
// `slides-core` maps a `PresentationImportResult` to a deck (`importFromPptx`, plan K15). The reader (plan K9) grows
// this result in steps: package, presentation and theme first, then layouts, then slide content.

import type { OpaquePart, OpaqueRef } from '../../opaque/types';

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
  /** The `p:sp` shapes, `p:pic` pictures, table `p:graphicFrame`s, `p:cxnSp` connectors and `p:grpSp` groups in drawing order; the other shape kinds arrive with the later reader steps. Absent for a slide that failed. */
  elements?: PptxElement[];
  /** The speaker notes (plan K10c): the body placeholder of the slide's notes page, its text styles inherited from the notes master. Absent when the slide has no notes text. */
  notes?: PptxTextBody;
  /** The comment threads of the slide (plan K10c), legacy and modern, in file order. Absent when it has none. */
  threads?: PptxCommentThread[];
  /** The slide's transition and animation XML (plan K11), kept as written (see `PptxSlideOpaque`). Absent when it has neither. */
  opaque?: PptxSlideOpaque;
  error?: string;
}

/**
 * Slide-level XML the editor does not model, as self-contained fragments (the namespaces the slide root declares are repeated on
 * the fragment's own tag): `transition` is the `p:transition` element, or the `mc:AlternateContent` that wraps it (a `p14`
 * transition with its fallback), `timing` is the `p:timing` element (animations and media playback). `timing` names shapes by
 * `pptxId` (`spTgt spid`); the editor prunes those references when it deletes the shape.
 */
export interface PptxSlideOpaque {
  transition?: string;
  timing?: string;
}

/** One comment of a thread: the author as the file names them (`Unknown author` when its list does not), the plain text of the body (paragraphs joined by a line break) and the time as UTC ISO-8601 when the file states one. */
export interface PptxComment {
  author: string;
  initials?: string;
  body: string;
  createdAt?: string;
}

/**
 * A comment thread of a slide (plan K10c). `format` says which file format it came from. A legacy comment (`p:cm`) and its
 * replies (a `cm` with `p15:parentCm`) and a modern one (`p188:cm` with `p188:reply`) both become `comments` with the
 * root first. `anchor` is the slide, or the element whose `pptxId` it is (modern comments only; an id the slide does not hold anchors to the slide).
 * `position` is a legacy comment's marker position on the slide in px.
 */
export interface PptxCommentThread {
  id: string;
  format: 'legacy' | 'modern';
  anchor: { kind: 'slide' } | { kind: 'element'; pptxId: string };
  position?: { x: number; y: number };
  resolved?: boolean;
  comments: PptxComment[];
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
 * One drawing command of a custom path (`a:moveTo`, `a:lnTo`, `a:cubicBezTo`, `a:quadBezTo`, `a:arcTo`, `a:close`), in the
 * path's own space (`PptxCustomPath.w` x `h`). `arc` angles are degrees clockwise from the +x axis (`arcTo` angles are 60000ths).
 */
export type PptxPathCommand =
  | { op: 'move'; x: number; y: number }
  | { op: 'line'; x: number; y: number }
  | { op: 'cubic'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: 'quad'; x1: number; y1: number; x: number; y: number }
  | { op: 'arc'; wR: number; hR: number; stAng: number; swAng: number }
  | { op: 'close' };

/** One `a:path` of custom geometry: commands in a `w` x `h` space (EMU unless the file gives the path a size) scaled onto the shape's box. */
export interface PptxCustomPath {
  w: number;
  h: number;
  /** `a:path/@fill` when not `norm`: `none`, `lighten`, `lightenLess`, `darken` or `darkenLess`. */
  fill?: string;
  /** `stroke="0"`: the path is not outlined. */
  stroke?: false;
  extrusionOk?: true;
  commands: PptxPathCommand[];
}

/**
 * A shape's geometry (plan K5): a preset by its OOXML `prst` name with its adjustments (`known` is false for a preset
 * the editor's library lacks, which the editor keeps under that name and draws approximately), or custom geometry (plan K10b).
 * `paths` of custom geometry is absent when the geometry cannot be kept faithfully (a coordinate that is a guide formula, a
 * malformed command, no path, too large): the shape then keeps its frame, fill and text without its outline.
 */
export type PptxGeometry =
  | { kind: 'preset'; prst: string; adjustments?: Record<string, number>; known: boolean }
  | { kind: 'custom'; paths?: PptxCustomPath[] };

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
 * in `part`). An SVG with a PNG fallback reads as the PNG. An SVG `dataUrl` is the file as it is, unsanitised: show it only
 * through an `<img>` and never inline it into the page.
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
  /** Widths in px. The grid is as wide as the widest row: a column the `a:tblGrid` does not list has width 0. */
  columns: number[];
  rows: { height: number; cells: PptxTableCell[] }[];
  /** The `a:tblPr` switches that are on. */
  flags?: Partial<Record<'firstRow' | 'firstCol' | 'lastRow' | 'lastCol' | 'bandRow' | 'bandCol', true>>;
  /** `a:tableStyleId` and the style's name, when the file defines it. */
  styleId?: string;
  styleName?: string;
}

/** The shape a connector end is glued to (`a:stCxn`, `a:endCxn`): the `pptxId` of the shape and the index of its connection site. */
export interface PptxConnection {
  id: string;
  idx: number;
}

/**
 * A connector (`p:cxnSp`, plan K10b): a line, bent (elbow) or curved route in `geometry` (`straightConnector1`, `bentConnector3`,
 * `curvedConnector2`... with their adjustments), the outline with its arrow heads (`line.head`, `line.tail`) and the shapes its ends are glued to.
 * `frame` is the box of the two ends with the flips that say which corners they are.
 */
export interface PptxConnector {
  kind: 'connector';
  pptxId: string;
  name: string;
  altText?: string;
  hidden?: boolean;
  frame?: PptxFrame;
  geometry?: PptxGeometry;
  line?: PptxLine;
  start?: PptxConnection;
  end?: PptxConnection;
}

/**
 * A group (`p:grpSp`, plan K10b). `frame` is the group's box in slide space (after the groups it sits in). The frames of everything
 * in `children` (nested groups included) are **composed**: they are already in slide space, so a child needs no group
 * transform; the group's own `a:chOff` / `a:chExt` mapping has been applied to them.
 */
export interface PptxGroup {
  kind: 'group';
  pptxId: string;
  name: string;
  altText?: string;
  hidden?: boolean;
  frame?: PptxFrame;
  children: PptxElement[];
}

/** What an opaque element is: a chart, a SmartArt diagram, an OLE object (or control), an audio or video clip, or ink. */
export type PptxOpaqueKind = 'chart' | 'smartArt' | 'ole' | 'media' | 'ink';

/**
 * Content the editor cannot edit or draw (plan K11, decision D5): kept byte for byte in the opaque part store
 * (`PresentationImportResult.opaqueParts`) and drawn as a labelled placeholder. `parts` refer to the stored parts: the root
 * part first (the chart, the diagram data, the OLE object, the media file, the ink), then everything it relates to
 * (embedded workbook, style and colour parts, diagram layout, drawing, preview and poster images), each with the package part
 * name it had and a `role` (the last segment of the relationship type: `chart`, `package`, `diagramData`, `image`, `media`...).
 * The relationship part of every stored part is stored too (role `rels`) so an exporter can rewrite its targets.
 * A part the store refused (over the size cap) is not listed; the importer warns. `props` holds what is needed to write the frame
 * again (`progId` and `name` of an OLE object, the `c:chart` / `dgm:relIds` / `p:oleObj` / media wrapper kind).
 */
export interface PptxOpaque {
  kind: 'opaque';
  opaqueKind: PptxOpaqueKind;
  pptxId: string;
  name: string;
  altText?: string;
  hidden?: boolean;
  frame?: PptxFrame;
  /** The chart title when the chart states one as plain text, for the placeholder's label. */
  title?: string;
  parts: OpaqueRef[];
  /** A picture to draw in the placeholder (an OLE object's preview, a clip's poster frame), as a data URL; never fetched. */
  poster?: string;
  props?: Record<string, string>;
}

/** What a slide holds, in drawing order. */
export type PptxElement = PptxShape | PptxPicture | PptxTable | PptxConnector | PptxGroup | PptxOpaque;

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
  /** The parts the opaque elements refer to (plan K11), content-addressed: one entry per distinct byte content. */
  opaqueParts: OpaquePart[];
  warnings: string[];
}
