// Scene layer types — the container-agnostic geometry + element model.
//
// Elements carry a `containerId` (a slide id today; a canvas/board id when a
// future spindle-canvas reuses this layer) but never any camera/viewport
// state. Everything here is plain data: the engine stores immutable records
// and replaces them on mutation so React snapshots version by reference.

import type { RichTextDoc } from '../text/model';

/**
 * A positioned box. `rotation` is in degrees, clockwise, about the box centre.
 * x/y are the top-left of the un-rotated box in slide coordinates (px).
 */
export interface Frame {
  x: number;
  y: number;
  w: number;
  h: number;
  rotation: number;
}

// ── Theme references ────────────────────────────────────────────────────────

/**
 * The 12 symbolic theme color slots (PPTX-compatible). Elements reference a
 * slot rather than a literal hex, so a theme switch recolors the whole deck.
 */
export type ThemeColorSlot =
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
 * A color is either a reference to a theme slot (with optional PPTX luminance
 * transforms) or a literal RGB. `alpha` is 0–1; omitted means fully opaque.
 */
export type Color =
  | { kind: 'theme'; slot: ThemeColorSlot; alpha?: number; lumMod?: number; lumOff?: number }
  | { kind: 'rgb'; hex: string; alpha?: number };

/** One colour stop of a gradient; `pos` is 0–1 along the gradient, the colour's `alpha` gives per-stop transparency. */
export interface GradientStop {
  pos: number;
  color: Color;
}

/** A gradient fill. `linear` runs along `angle` (degrees, clockwise from left-to-right, as in DrawingML `a:lin`);
 *  `radial` is an ellipse from the box centre; `path` is the DrawingML path gradient, stop 0 at the centre. */
export interface GradientFill {
  kind: 'gradient';
  type: 'linear' | 'radial' | 'path';
  /** Degrees; only meaningful for `linear` (default 90, top to bottom is 90). */
  angle?: number;
  stops: GradientStop[];
}

/** A picture fill. `src` is a data URL (no network fetch ever); `tile` is the tile size in px when `mode` is `tile`. */
export interface PictureFill {
  kind: 'picture';
  src: string;
  mode: 'stretch' | 'tile';
  tile?: { w: number; h: number };
  /** 0–1; omitted means opaque. */
  alpha?: number;
}

/** A DrawingML pattern fill (`prst` is the preset name, e.g. `pct50`, `dkHorz`). Rendered as the solid blend of
 *  its two colours (approximated; the pattern itself is kept in the model). */
export interface PatternFill {
  kind: 'pattern';
  preset: string;
  fg: Color;
  bg: Color;
}

export type Fill =
  | { kind: 'none' }
  | { kind: 'solid'; color: Color }
  | GradientFill
  | PictureFill
  | PatternFill;

export type StrokeDash = 'solid' | 'dash' | 'dot';

export interface Stroke {
  color: Color;
  width: number;
  dash?: StrokeDash;
}

// ── Placeholder metadata (PPTX fidelity, in the schema from day one) ─────────

export type PlaceholderType =
  | 'title'
  | 'centerTitle'
  | 'subtitle'
  | 'body'
  | 'image';

export interface PlaceholderMeta {
  type: PlaceholderType;
  idx: number;
}

// ── Text body styling (element-level; paragraph/run styling lives in PM) ─────

/** Inner padding per side in px (`a:bodyPr` `lIns`/`tIns`/`rIns`/`bIns`, converted from EMU). */
export interface BodyInsets {
  l: number;
  t: number;
  r: number;
  b: number;
}

/**
 * What a text body does when its text does not fit the box (`a:noAutofit`, `a:normAutofit`, `a:spAutoFit`).
 * `shrink` stores the scales the file (or the last edit) computed and applies them as stored: `fontScale` multiplies
 * every font size (0.625 is 62.5%) and `lnSpcReduction` takes that share off the line spacing (0.2 is 20%). They are
 * recomputed only when the user edits the text, as PowerPoint does. `resize` grows or shrinks the box height to the text.
 */
export interface Autofit {
  mode: 'none' | 'shrink' | 'resize';
  /** 0.1-1; omitted means 1 (the text fits unscaled). Only meaningful for `shrink`. */
  fontScale?: number;
  /** 0-0.5; omitted means 0. Only meaningful for `shrink`. */
  lnSpcReduction?: number;
}

/** Text direction of a body (`a:bodyPr/@vert`): `vert` rotates horizontal text 90 degrees, `eaVert` stacks East Asian text upright. */
export type TextVert = 'horz' | 'vert' | 'eaVert';

export interface BodyStyle {
  vAlign?: 'top' | 'middle' | 'bottom';
  /** Uniform inner padding in px (v1). `insets`, when present, wins. */
  padding?: number;
  /** Inner padding per side in px. */
  insets?: BodyInsets;
  /** Centre the text block horizontally in the box (`anchorCtr`), independent of paragraph alignment. */
  anchorCtr?: boolean;
  autofit?: Autofit;
  vert?: TextVert;
  /** Text columns (`numCol`, `spcCol`); `gap` in px. */
  columns?: { count: number; gap?: number };
  /** Wrap text to the box width (default true). */
  wrap?: boolean;
  // Element-level text defaults (PPTX defRPr analogue): runs that don't carry
  // an explicit mark inherit these from the container, so placeholder prompt
  // text and freshly-typed text both render at the intended size/color.
  fontSize?: number;
  color?: Color;
  bold?: boolean;
  /** 'major' | 'minor' | a literal family. */
  fontFamily?: string;
}

// ── Elements ─────────────────────────────────────────────────────────────────

export interface ElementBase extends Frame {
  id: string;
  /** The slide (or, later, board) this element belongs to. */
  containerId: string;
  /** Fractional z-order key — sorted ascending, painted back-to-front. */
  index: string;
  /** The innermost group this element belongs to; members of a group share its id. */
  groupId?: string;
  /** The group that contains `groupId` (nested groups, K6). Every member of a nested group carries the same value; the
   *  full chain is read by following it (`groupChain`). Absent for a top-level group. */
  parentGroupId?: string;
  locked?: boolean;
  /** 0–1; omitted means 1 (fully opaque). */
  opacity?: number;
  /** Present when this element fills a layout placeholder. */
  placeholder?: PlaceholderMeta;
}

/**
 * The built-in shape presets (K5): the original 18 plus 58 more (76 in all), named after the OOXML `prst` values so an
 * imported preset keeps its name. Geometry lives in `slides-react` (`shapeGeom`); `shape-presets.ts` holds the metadata.
 */
export type ShapePreset =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'rightTriangle'
  | 'diamond'
  | 'pentagon'
  | 'hexagon'
  | 'octagon'
  | 'star5'
  | 'arrowRight'
  | 'arrowLeft'
  | 'chevron'
  | 'parallelogram'
  | 'trapezoid'
  | 'plus'
  | 'heart'
  | 'cloud'
  // rectangles
  | 'snip1Rect'
  | 'snip2SameRect'
  | 'snip2DiagRect'
  | 'snipRoundRect'
  | 'round1Rect'
  | 'round2SameRect'
  | 'round2DiagRect'
  | 'foldedCorner'
  | 'frame'
  | 'plaque'
  | 'bevel'
  | 'can'
  | 'cube'
  // basic
  | 'donut'
  | 'noSmoking'
  | 'teardrop'
  | 'pie'
  | 'chord'
  | 'blockArc'
  | 'moon'
  | 'smileyFace'
  | 'lightningBolt'
  | 'heptagon'
  | 'decagon'
  | 'dodecagon'
  // block arrows
  | 'rightArrow'
  | 'leftArrow'
  | 'upArrow'
  | 'downArrow'
  | 'leftRightArrow'
  | 'upDownArrow'
  | 'homePlate'
  // stars
  | 'star4'
  | 'star6'
  | 'star8'
  | 'star12'
  // callouts
  | 'wedgeRectCallout'
  | 'wedgeRoundRectCallout'
  | 'wedgeEllipseCallout'
  // flowchart
  | 'flowChartProcess'
  | 'flowChartAlternateProcess'
  | 'flowChartDecision'
  | 'flowChartInputOutput'
  | 'flowChartTerminator'
  | 'flowChartConnector'
  | 'flowChartPredefinedProcess'
  | 'flowChartDelay'
  | 'flowChartPreparation'
  | 'flowChartOffpageConnector'
  | 'flowChartMerge'
  // brackets and braces
  | 'leftBracket'
  | 'rightBracket'
  | 'leftBrace'
  | 'rightBrace'
  | 'bracketPair'
  // math and line inversion
  | 'mathPlus'
  | 'mathMinus'
  | 'lineInv';

/**
 * A preset the editor has no geometry for, kept under its original OOXML name: `unknown:<prst>`. It is drawn as a
 * rounded rectangle approximation and re-exported verbatim, so the geometry is never lost, only drawn approximately.
 */
export type UnknownShape = `unknown:${string}`;

/** The shape kind of a custom geometry (`a:custGeom`): the outline is `ShapeElement.path`, not a preset. */
export type PathShape = 'path';

/** What `ShapeElement.shape` holds: a built-in preset, an unknown preset passed through, or `path` (custom geometry). */
export type ShapeKind = ShapePreset | UnknownShape | PathShape;

/** One drawing command of a custom path. Coordinates are in the path's own space (`CustomPath.w` x `CustomPath.h`);
 *  `arc` angles are degrees clockwise from the +x axis (OOXML `arcTo`, converted from 60000ths) and the arc starts at
 *  the current point. */
export type PathCommand =
  | { op: 'move'; x: number; y: number }
  | { op: 'line'; x: number; y: number }
  | { op: 'cubic'; x1: number; y1: number; x2: number; y2: number; x: number; y: number }
  | { op: 'quad'; x1: number; y1: number; x: number; y: number }
  | { op: 'arc'; wR: number; hR: number; stAng: number; swAng: number }
  | { op: 'close' };

/** How a path is filled (`a:path/@fill`): `none` leaves it unfilled; the others shade the shape's fill. */
export type PathFillMode = 'norm' | 'none' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';

/** One `a:path`: commands in a `w` x `h` coordinate space that is scaled onto the shape's box. */
export interface CustomPath {
  w: number;
  h: number;
  /** Omitted means `norm`. */
  fill?: PathFillMode;
  /** Omitted means true (the path is outlined). */
  stroke?: boolean;
  extrusionOk?: boolean;
  commands: PathCommand[];
}

/** Custom geometry (`a:custGeom`): one or more paths drawn over the same box. */
export interface CustomGeometry {
  paths: CustomPath[];
}

export type ArrowHead = 'none' | 'arrow' | 'triangle' | 'circle';

/** The 8 connection anchors on a box: 4 edge midpoints + 4 corners. */
export type AnchorId = 'n' | 'e' | 's' | 'w' | 'nw' | 'ne' | 'se' | 'sw';

/**
 * A connector endpoint bound to another element's anchor. The endpoint's pixel
 * position is derived from the target's current frame — never stored — so the
 * connector tracks the shape when it moves, resizes, or rotates.
 */
export interface EndpointBind {
  elementId: string;
  anchor: AnchorId;
}

export interface TextElement extends ElementBase {
  type: 'text';
  richText: RichTextDoc;
  bodyStyle?: BodyStyle;
  fill?: Fill;
  stroke?: Stroke;
}

export interface ShapeElement extends ElementBase {
  type: 'shape';
  shape: ShapeKind;
  /** The outline of a `path` shape (custom geometry). Ignored for any other kind. */
  path?: CustomGeometry;
  /**
   * Preset-specific adjustment values in the order of `SHAPE_PRESET_INFO[shape].adj`: OOXML `adj` guides divided by
   * 100000 (angles by 21600000), so 0.5 is 50000. Missing entries take the preset default. An `unknown:` shape keeps
   * the raw guide values in file order, uninterpreted.
   */
  adjustments?: number[];
  flipH?: boolean;
  flipV?: boolean;
  fill: Fill;
  stroke?: Stroke;
  richText?: RichTextDoc;
  bodyStyle?: BodyStyle;
}

export interface ImageElement extends ElementBase {
  type: 'image';
  src: string;
  naturalW: number;
  naturalH: number;
  flipH?: boolean;
  flipV?: boolean;
  /**
   * How the image sits in its box (CSS object-fit). 'fill' stretches to the box
   * (default, back-compat); 'contain' letterboxes; 'cover' crops to fill without
   * distortion. With a `crop`, the fit applies to the cropped region.
   */
  fit?: 'fill' | 'contain' | 'cover';
  /** Region crop: the share of the source image trimmed from each edge. Omitted or all zero means uncropped. */
  crop?: ImageCrop;
}

/** Percent (0-100) of the source image's width (l, r) or height (t, b) trimmed from each edge (`a:srcRect`). */
export interface ImageCrop {
  l: number;
  t: number;
  r: number;
  b: number;
}

/** A frame plus the preset and flips of a shape element, so connection sites can follow its outline (see `anchorPoint`). */
export interface ConnectionFrame extends Frame {
  shape?: ShapeKind;
  /** Mirrored like the rendered element: sites and exit directions are flipped in the box before rotating. */
  flipH?: boolean;
  flipV?: boolean;
}

/** How a connector is routed between its endpoints: a straight segment, right-angle (elbow) legs or an S-curve. */
export type ConnectorKind = 'straight' | 'elbow' | 'curved';

export interface LineElement extends ElementBase {
  type: 'line';
  /** Routing of a connector (default `straight`). Elbow and curved routes leave and enter bound anchors perpendicular to their side. */
  connector?: ConnectorKind;
  /** Elbow only: where the middle leg sits between the two ends, 0 to 1 (OOXML `bentConnector3` adj1; default 0.5). */
  adjustments?: number[];
  /** Diagonal orientation flips (a line occupies the box diagonal). */
  flipH?: boolean;
  flipV?: boolean;
  stroke: Stroke;
  startArrow?: ArrowHead;
  endArrow?: ArrowHead;
  /** When set, the endpoint tracks this element's anchor instead of the box corner. */
  startBind?: EndpointBind;
  endBind?: EndpointBind;
  /**
   * Explicit free-endpoint positions (slide coords), used when the endpoint is
   * a connector end that is NOT bound — the box + flipV alone can't say which
   * corner a free end occupies once the other end is bound, so we store it.
   */
  startPoint?: { x: number; y: number };
  endPoint?: { x: number; y: number };
}

/** One cell of a table: its own rich-text body, optional fill and vertical align. */
export interface TableCell {
  richText: RichTextDoc;
  fill?: Fill;
  bodyStyle?: BodyStyle;
}

export interface TableElement extends ElementBase {
  type: 'table';
  rows: number;
  cols: number;
  /** Column widths as fractions of the table width (sum ≈ 1). */
  colFractions: number[];
  /** Row heights as fractions of the table height (sum ≈ 1). Legacy sizing; rows
   *  are now content-driven with an optional per-row minimum (rowHeights). */
  rowFractions: number[];
  /** Optional per-row minimum height in px (0/absent = auto). A row grows to fit
   *  its content and can be dragged taller than the content via this minimum. */
  rowHeights?: number[];
  /** Row-major grid: cells[row][col]. */
  cells: TableCell[][];
  /** Grid line style (applied to every cell edge). */
  border?: Stroke;
}

export type SlideElement = TextElement | ShapeElement | ImageElement | LineElement | TableElement;
export type ElementType = SlideElement['type'];
