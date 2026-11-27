import type { PptxColorMap, PptxLayout, PptxPlaceholder } from '@b32nio/spindle-shared/export/pptx';
import { bodyPrXml, lstStyleXml, plainParagraphsXml } from './writer-text';
import { attrs, esc, ROOT_NS, solidFillXml, TREE_HEADER, XML_DECLARATION, xfrmXml } from './writer-xml';

// The slide master and the slide layouts (plan K12). The deck keeps one theme and the importer flattens master -> layout ->
// slide inheritance (decision SD6), so the master written here is a plain one (the five standard placeholders, empty text
// styles) and every layout states its placeholders in full: position, text levels and body properties. Reading the result
// gives the layouts back; the master adds nothing the layouts did not say.

/** `ST_SlideLayoutType`: the values a layout's `type` may take. */
const LAYOUT_TYPES = new Set([
  'title', 'tx', 'twoColTx', 'tbl', 'txAndChart', 'chartAndTx', 'dgm', 'chart', 'txAndClipArt', 'clipArtAndTx', 'titleOnly', 'blank',
  'txAndObj', 'objAndTx', 'objOnly', 'obj', 'txAndMedia', 'mediaAndTx', 'objOverTx', 'txOverObj', 'txAndTwoObj', 'twoObjAndTx',
  'twoObjOverTx', 'fourObj', 'vertTx', 'clipArtAndVertTx', 'vertTitleAndTx', 'vertTitleAndTxOverChart', 'twoObj', 'objAndTwoObj',
  'twoObjAndObj', 'cust', 'secHead', 'twoTxTwoObj', 'objTx', 'picTx',
]);

/** The master's first layout id: `p:sldMasterId` takes 2147483648 and the layouts follow it, as ids of both lists are unique together. */
export const FIRST_LAYOUT_ID = 2147483649;
export const MASTER_ID = 2147483648;

const COLOR_MAP_ROLES: (keyof PptxColorMap)[] = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** A layout to write: the model's layout, its part name and its place in the master's list. */
export interface LayoutPart {
  layout: PptxLayout;
  part: string;
}

/** Layout used when the deck has none (a presentation needs at least one) or a slide names one that does not exist. */
export const BLANK_LAYOUT: PptxLayout = { partName: '', name: 'Blank', layoutType: 'blank', masterPart: '', placeholders: [], elements: [], staticShapes: 0 };

interface ShapeInfo {
  id: number;
  name: string;
  /** `p:ph` attributes. */
  ph: Record<string, string | number | undefined>;
  frame?: PptxPlaceholder['frame'];
  body: string;
}

function placeholderXml(s: ShapeInfo): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr${attrs({ id: s.id, name: s.name })}/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${attrs(s.ph)}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${s.frame ? xfrmXml(s.frame) : ''}</p:spPr><p:txBody>${s.body}</p:txBody></p:sp>`
  );
}

/** A layout placeholder as the model states it: its type and `idx`, its frame, body properties, text levels and prompt. */
function layoutPlaceholder(p: PptxPlaceholder, id: number): string {
  return placeholderXml({
    id,
    name: p.name ?? `${p.type} ${id - 1}`,
    ph: { type: p.type === 'obj' ? undefined : p.type, idx: p.idx },
    frame: p.frame,
    body: `${bodyPrXml(p.bodyProps, p.anchor)}${lstStyleXml(p.levels)}${plainParagraphsXml(p.prompt ?? '')}`,
  });
}

export function layoutXml(layout: PptxLayout): string {
  const { layoutType } = layout;
  const background = layout.background ? `<p:bg><p:bgPr>${solidFillXml(layout.background)}<a:effectLst/></p:bgPr></p:bg>` : '';
  return (
    `${XML_DECLARATION}<p:sldLayout ${ROOT_NS}${attrs({ type: layoutType && LAYOUT_TYPES.has(layoutType) ? layoutType : undefined, showMasterSp: layout.hideMasterShapes ? 0 : undefined, preserve: 1 })}>` +
    `<p:cSld name="${esc(layout.name)}">${background}<p:spTree>${TREE_HEADER}${layout.placeholders.map((p, i) => layoutPlaceholder(p, i + 2)).join('')}</p:spTree></p:cSld>` +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
  );
}

/** The master's standard placeholders, placed on the slide the way Office's default master does (as fractions of the slide). */
function masterPlaceholders(size: { w: number; h: number }): string {
  const box = (x: number, y: number, w: number, h: number) => ({ x: size.w * x, y: size.h * y, w: size.w * w, h: size.h * h });
  const shapes: ShapeInfo[] = [
    { id: 2, name: 'Title Placeholder 1', ph: { type: 'title' }, frame: box(0.0688, 0.0486, 0.8625, 0.1764), body: `<a:bodyPr/><a:lstStyle/>${plainParagraphsXml('Click to edit Master title style')}` },
    { id: 3, name: 'Text Placeholder 2', ph: { type: 'body', idx: 1 }, frame: box(0.0688, 0.2431, 0.8625, 0.5792), body: `<a:bodyPr/><a:lstStyle/>${plainParagraphsXml('Click to edit Master text styles')}` },
    { id: 4, name: 'Date Placeholder 3', ph: { type: 'dt', sz: 'half', idx: 2 }, frame: box(0.0688, 0.9028, 0.2813, 0.0486), body: `<a:bodyPr/><a:lstStyle/>${plainParagraphsXml('')}` },
    { id: 5, name: 'Footer Placeholder 4', ph: { type: 'ftr', sz: 'quarter', idx: 3 }, frame: box(0.3438, 0.9028, 0.3125, 0.0486), body: `<a:bodyPr/><a:lstStyle/>${plainParagraphsXml('')}` },
    { id: 6, name: 'Slide Number Placeholder 5', ph: { type: 'sldNum', sz: 'quarter', idx: 4 }, frame: box(0.6500, 0.9028, 0.2813, 0.0486), body: `<a:bodyPr/><a:lstStyle/>${plainParagraphsXml('')}` },
  ];
  return shapes.map(placeholderXml).join('');
}

/** `ppt/slideMasters/slideMaster1.xml`: placeholders, the colour map, the list of layouts (`rIds` in order) and empty text styles. */
export function masterXml(size: { w: number; h: number }, colorMap: PptxColorMap, layoutRelIds: string[], name?: string): string {
  const map = attrs(Object.fromEntries(COLOR_MAP_ROLES.map((r) => [r, colorMap[r]])));
  const ids = layoutRelIds.map((rid, i) => `<p:sldLayoutId id="${FIRST_LAYOUT_ID + i}" r:id="${rid}"/>`).join('');
  return (
    `${XML_DECLARATION}<p:sldMaster ${ROOT_NS}><p:cSld${attrs({ name })}><p:spTree>${TREE_HEADER}${masterPlaceholders(size)}</p:spTree></p:cSld>` +
    `<p:clrMap${map}/><p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst>` +
    '<p:txStyles><p:titleStyle><a:lvl1pPr/></p:titleStyle><p:bodyStyle><a:lvl1pPr/></p:bodyStyle><p:otherStyle><a:lvl1pPr/></p:otherStyle></p:txStyles></p:sldMaster>'
  );
}
