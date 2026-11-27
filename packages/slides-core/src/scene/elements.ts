// Element factories — build fully-formed, defaulted element records. The deck
// engine supplies `containerId` and a computed z-order `index`; callers supply
// the interesting bits. `id` may be provided (setData / duplicate) or is
// generated fresh.

import { generateId } from '../utils/id';
import { emptyRichText } from '../text/model';
import { normalizeCrop } from './crop';
import type {
  Frame,
  Fill,
  Stroke,
  BodyStyle,
  ElementBase,
  TextElement,
  ShapeElement,
  ImageElement,
  ImageCrop,
  LineElement,
  ConnectorKind,
  TableElement,
  TableCell,
  TableStyleFlags,
  ShapeKind,
  ArrowHead,
  EndpointBind,
  PlaceholderMeta,
  ElementType,
} from './types';
import { evenFractions, makeGrid, normalizeSpans } from './table';

/** Fields every element needs beyond its type-specific payload. */
export interface BaseElementInput extends Partial<Frame> {
  containerId: string;
  index: string;
  id?: string;
  groupId?: string;
  parentGroupId?: string;
  locked?: boolean;
  opacity?: number;
  placeholder?: PlaceholderMeta;
}

const DEFAULT_FRAME: Record<ElementType, Frame> = {
  text: { x: 100, y: 100, w: 400, h: 100, rotation: 0 },
  shape: { x: 100, y: 100, w: 240, h: 200, rotation: 0 },
  image: { x: 100, y: 100, w: 320, h: 240, rotation: 0 },
  line: { x: 100, y: 100, w: 300, h: 0, rotation: 0 },
  table: { x: 120, y: 160, w: 720, h: 300, rotation: 0 },
};

const DEFAULT_BODY_STYLE: BodyStyle = { vAlign: 'top', padding: 8, wrap: true };

function base(input: BaseElementInput, type: ElementType): ElementBase {
  const f = DEFAULT_FRAME[type];
  const el: ElementBase = {
    id: input.id ?? generateId(),
    containerId: input.containerId,
    index: input.index,
    x: input.x ?? f.x,
    y: input.y ?? f.y,
    w: input.w ?? f.w,
    h: input.h ?? f.h,
    rotation: input.rotation ?? f.rotation,
  };
  if (input.groupId !== undefined) el.groupId = input.groupId;
  if (input.parentGroupId !== undefined) el.parentGroupId = input.parentGroupId;
  if (input.locked !== undefined) el.locked = input.locked;
  if (input.opacity !== undefined) el.opacity = input.opacity;
  if (input.placeholder !== undefined) el.placeholder = input.placeholder;
  return el;
}

export interface TextElementInput extends BaseElementInput {
  richText?: TextElement['richText'];
  bodyStyle?: BodyStyle;
  fill?: Fill;
  stroke?: Stroke;
}

export function createTextElement(input: TextElementInput): TextElement {
  return {
    ...base(input, 'text'),
    type: 'text',
    richText: input.richText ?? emptyRichText(),
    bodyStyle: input.bodyStyle ?? { ...DEFAULT_BODY_STYLE },
    ...(input.fill ? { fill: input.fill } : {}),
    ...(input.stroke ? { stroke: input.stroke } : {}),
  };
}

export interface ShapeElementInput extends BaseElementInput {
  shape?: ShapeKind;
  path?: ShapeElement['path'];
  adjustments?: number[];
  flipH?: boolean;
  flipV?: boolean;
  fill?: Fill;
  stroke?: Stroke;
  richText?: ShapeElement['richText'];
  bodyStyle?: BodyStyle;
}

export function createShapeElement(input: ShapeElementInput): ShapeElement {
  const el: ShapeElement = {
    ...base(input, 'shape'),
    type: 'shape',
    shape: input.shape ?? 'rect',
    fill: input.fill ?? { kind: 'solid', color: { kind: 'theme', slot: 'accent1' } },
  };
  if (input.path) el.path = input.path;
  if (input.adjustments) el.adjustments = input.adjustments;
  if (input.flipH) el.flipH = input.flipH;
  if (input.flipV) el.flipV = input.flipV;
  if (input.stroke) el.stroke = input.stroke;
  if (input.richText) el.richText = input.richText;
  if (input.bodyStyle) el.bodyStyle = input.bodyStyle;
  return el;
}

export interface ImageElementInput extends BaseElementInput {
  src: string;
  naturalW: number;
  naturalH: number;
  flipH?: boolean;
  flipV?: boolean;
  fit?: 'fill' | 'contain' | 'cover';
  crop?: ImageCrop;
}

export function createImageElement(input: ImageElementInput): ImageElement {
  const el: ImageElement = {
    ...base(input, 'image'),
    type: 'image',
    src: input.src,
    naturalW: input.naturalW,
    naturalH: input.naturalH,
  };
  if (input.flipH) el.flipH = input.flipH;
  if (input.flipV) el.flipV = input.flipV;
  if (input.fit) el.fit = input.fit;
  const crop = normalizeCrop(input.crop);
  if (crop) el.crop = crop;
  return el;
}

export interface LineElementInput extends BaseElementInput {
  stroke?: Stroke;
  startArrow?: ArrowHead;
  endArrow?: ArrowHead;
  flipH?: boolean;
  flipV?: boolean;
  startBind?: EndpointBind;
  endBind?: EndpointBind;
  startPoint?: { x: number; y: number };
  endPoint?: { x: number; y: number };
  connector?: ConnectorKind;
  adjustments?: number[];
}

export function createLineElement(input: LineElementInput): LineElement {
  const el: LineElement = {
    ...base(input, 'line'),
    type: 'line',
    stroke: input.stroke ?? { color: { kind: 'theme', slot: 'dk1' }, width: 2 },
  };
  if (input.startArrow) el.startArrow = input.startArrow;
  if (input.endArrow) el.endArrow = input.endArrow;
  if (input.flipH) el.flipH = input.flipH;
  if (input.flipV) el.flipV = input.flipV;
  if (input.startBind) el.startBind = input.startBind;
  if (input.endBind) el.endBind = input.endBind;
  if (input.startPoint) el.startPoint = input.startPoint;
  if (input.endPoint) el.endPoint = input.endPoint;
  if (input.connector && input.connector !== 'straight') el.connector = input.connector;
  if (input.adjustments?.length) el.adjustments = input.adjustments;
  return el;
}

export interface TableElementInput extends BaseElementInput {
  rows: number;
  cols: number;
  colFractions?: number[];
  rowFractions?: number[];
  cells?: TableCell[][];
  border?: Stroke;
  styleFlags?: TableStyleFlags;
}

export function createTableElement(input: TableElementInput): TableElement {
  const rows = Math.max(1, input.rows);
  const cols = Math.max(1, input.cols);
  // Normalize cells to exactly rows×cols so malformed authored JSON (a cells
  // grid that doesn't match rows/cols) can't throw on cells[r][c] at render.
  const grid = makeGrid(rows, cols);
  if (input.cells) {
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (input.cells[r]?.[c]) grid[r][c] = input.cells[r][c];
      }
    }
  }
  const el: TableElement = {
    ...base(input, 'table'),
    type: 'table',
    rows,
    cols,
    colFractions: input.colFractions?.length === cols ? input.colFractions : evenFractions(cols),
    rowFractions: input.rowFractions?.length === rows ? input.rowFractions : evenFractions(rows),
    cells: normalizeSpans(grid, rows, cols),
    border: input.border ?? { color: { kind: 'theme', slot: 'dk2' }, width: 1 },
  };
  if (input.styleFlags) el.styleFlags = input.styleFlags;
  return el;
}

/** The default frame for a freshly-inserted element of a given type. */
export function defaultFrameFor(type: ElementType): Frame {
  return { ...DEFAULT_FRAME[type] };
}
