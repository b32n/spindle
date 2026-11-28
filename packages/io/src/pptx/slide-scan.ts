import type { Rel } from '../common/ooxml-package';
import { descendants, isOn, isTrue, kid, type XmlElement } from '../xlsx/xml';
import { DEFAULT_COLOR_MAP } from './color';
import { parseShadow } from './shadow';
import type { Tallies } from './tally';

// ---------------------------------------------------------------------------
// One pass over a slide part: every object the detector counts is found while walking the element tree once.
// Counting rules (what one "object" is) are documented per id in documentation/import-loss.md.
// ---------------------------------------------------------------------------

/**
 * The shape presets of slides-core (`ShapePreset`, K5), by OOXML `prst` name (`rtTriangle` is `rightTriangle` there), plus the two line geometries that become
 * line elements. A preset outside this set is kept as an `unknown:<prst>` shape and drawn approximately.
 */
export const SUPPORTED_PRESETS: ReadonlySet<string> = new Set([
  'rect', 'roundRect', 'ellipse', 'triangle', 'rightTriangle', 'rtTriangle', 'diamond', 'pentagon', 'hexagon', 'octagon', 'star5',
  'arrowRight', 'arrowLeft', 'chevron', 'parallelogram', 'trapezoid', 'plus', 'heart', 'cloud', 'snip1Rect',
  'snip2SameRect', 'snip2DiagRect', 'snipRoundRect', 'round1Rect', 'round2SameRect', 'round2DiagRect',
  'foldedCorner', 'frame', 'plaque', 'bevel', 'can', 'cube', 'donut', 'noSmoking', 'teardrop', 'pie', 'chord',
  'blockArc', 'moon', 'smileyFace', 'lightningBolt', 'heptagon', 'decagon', 'dodecagon', 'rightArrow', 'leftArrow',
  'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'homePlate', 'star4', 'star6', 'star8', 'star12',
  'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'flowChartProcess',
  'flowChartAlternateProcess', 'flowChartDecision', 'flowChartInputOutput', 'flowChartTerminator',
  'flowChartConnector', 'flowChartPredefinedProcess', 'flowChartDelay', 'flowChartPreparation',
  'flowChartOffpageConnector', 'flowChartMerge', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace',
  'bracketPair', 'mathPlus', 'mathMinus', 'lineInv',
  'line', 'straightConnector1',
]);

const ELBOW_CURVED = /^(?:bent|curved)Connector\d*$/;
const UNSUPPORTED_IMAGE = /\.(?:emf|wmf|emz|wmz)$/i;
const RECOLOR = new Set(['duotone', 'grayscl', 'biLevel', 'lum', 'clrChange', 'tint', 'hsl']);
const FILL_HOLDERS = new Set(['spPr', 'bgPr', 'tcPr', 'grpSpPr']);
const EFFECT_HOLDERS = new Set(['spPr', 'grpSpPr', 'rPr']);
const ZOOM = new Set(['sldZmObj', 'secZmObj', 'sldZm', 'secZm']);
/** Timing nodes that animate something; a tree with only media playback nodes (`cMediaNode`) is not an animation. */
const ANIMATES = new Set(['set', 'anim', 'animEffect', 'animMotion', 'animClr', 'animScale', 'animRot']);
const SHAPE_LIKE = new Set(['sp', 'pic', 'graphicFrame', 'cxnSp']);
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';

export interface SlideContext {
  /** `Slide 3`: where the findings are reported. */
  label: string;
  /** The slide's relationships by id (to resolve a picture's file). */
  rels: ReadonlyMap<string, Rel>;
  t: Tallies;
}

const nonZero = (v: string | undefined): boolean => v !== undefined && Number(v) !== 0;

/** The kind of an action hyperlink: `ppaction://hlinkshowjump?jump=nextslide` -> `hlinkshowjump`. */
function actionKind(action: string): string {
  const m = /^ppaction:\/\/([^?]*)/i.exec(action);
  return m ? m[1] : 'other';
}

export function scanSlide(root: XmlElement, c: SlideContext): void {
  const { t, label } = c;
  let groupDepth = 0;
  let transition = false;
  let timing = false;
  let effects = 0;
  let animates = false;
  const shadows = new Set<XmlElement>();

  const keptShadow = (list: XmlElement | undefined): void => {
    if (list && parseShadow(list, DEFAULT_COLOR_MAP)) shadows.add(list);
  };

  const geometry = (e: XmlElement): void => {
    const spPr = kid(e, 'spPr');
    if (kid(spPr, 'custGeom')) {
      t.custGeom.add(label);
      t.geometry.add('custGeom');
      return;
    }
    const prst = kid(spPr, 'prstGeom')?.attrs.prst;
    if (prst === undefined) return;
    if (ELBOW_CURVED.test(prst)) t.elbowCurved.add(label);
    else if (!SUPPORTED_PRESETS.has(prst)) t.geometry.add(prst);
  };

  const inheritsPosition = (e: XmlElement): boolean => {
    const nv = e.children.find((k) => k.name.startsWith('nv'));
    if (!kid(kid(nv, 'nvPr'), 'ph')) return false;
    return !(e.name === 'graphicFrame' ? kid(e, 'xfrm') : kid(kid(e, 'spPr'), 'xfrm'));
  };

  const picture = (e: XmlElement): void => {
    const rect = kid(e, 'srcRect');
    if (rect && ['l', 't', 'r', 'b'].some((a) => nonZero(rect.attrs[a]))) t.crop.add(label);
    const blip = kid(e, 'blip');
    if (!blip) return;
    if (blip.attrs.link) t.linked.add(label);
    if (descendants(blip, 'svgBlip').length > 0) t.svg.add(label);
    if (blip.children.some((k) => RECOLOR.has(k.name))) t.recolor.add(label);
    const rel = c.rels.get(blip.attrs.embed ?? blip.attrs.link ?? '');
    if (rel && UNSUPPORTED_IMAGE.test(rel.target)) t.unsupportedImages.add(label);
  };

  const paragraph = (e: XmlElement): void => {
    if (isTrue(e.attrs.rtl)) t.rtl.add(label);
    if (kid(e, 'buBlip')) t.bulletsPicture.add(label);
    const custom = e.children.some(
      (k) =>
        k.name === 'buClr' ||
        k.name === 'buSzPct' ||
        k.name === 'buSzPts' ||
        (k.name === 'buChar' && k.attrs.char !== '•') ||
        (k.name === 'buAutoNum' && (k.attrs.startAt !== undefined || k.attrs.type !== 'arabicPeriod'))
    );
    if (custom) t.bulletsCustom.add(label);
    if (kid(e, 'tabLst')?.children.some((k) => k.name === 'tab')) t.tabStops.add(label);
  };

  const textBody = (e: XmlElement): void => {
    const norm = kid(e, 'normAutofit');
    if ((norm && (norm.attrs.fontScale !== undefined || norm.attrs.lnSpcReduction !== undefined)) || kid(e, 'spAutoFit'))
      t.autofit.add(label);
    if (Number(e.attrs.numCol) >= 2) t.columns.add(label);
    if (e.attrs.vert !== undefined && e.attrs.vert !== 'horz') t.vertical.add(label);
    const warp = kid(e, 'prstTxWarp');
    if (warp && warp.attrs.prst !== 'textNoShape') t.wordArt.add(label);
  };

  const run = (e: XmlElement): void => {
    if (nonZero(e.attrs.spc) || e.attrs.cap === 'all' || e.attrs.cap === 'small') t.spacing.add(label);
    if (nonZero(e.attrs.baseline)) t.baseline.add(label);
    const lang = e.attrs.lang;
    if (lang && lang.toLowerCase() !== 'en-us') t.lang.add(lang);
  };

  const table = (e: XmlElement): void => {
    const pr = kid(e, 'tblPr');
    const a = pr?.attrs ?? {};
    if (isTrue(a.bandRow) || isTrue(a.bandCol)) t.banding.add(label);
    if (kid(pr, 'tableStyleId') || ['firstRow', 'firstCol', 'lastRow', 'lastCol'].some((k) => isTrue(a[k])))
      t.tableStyles.add(label);
  };

  const visit = (e: XmlElement, parent: XmlElement | undefined): void => {
    if (SHAPE_LIKE.has(e.name) && inheritsPosition(e)) t.layoutInheritance.add(label);
    if (e.children.some((k) => k.name === 'scene3d' || k.name === 'sp3d')) t.scene3d.add(label);
    switch (e.name) {
      case 'sp':
        // the reader keeps the shadow of a shape whose effect list holds only that
        keptShadow(kid(kid(e, 'spPr'), 'effectLst'));
        geometry(e);
        break;
      case 'cxnSp':
        geometry(e);
        break;
      case 'grpSp':
        if (groupDepth >= 1) t.nestedGroups.add(label);
        break;
      case 'graphicData': {
        const uri = e.attrs.uri ?? '';
        if (/chart/i.test(uri)) t.charts.add(label);
        else if (/\/diagram/i.test(uri)) t.smartArt.add(label);
        break;
      }
      case 'oleObj':
      case 'control':
        t.embedded.add(label);
        break;
      case 'videoFile':
      case 'audioFile':
      case 'quickTimeFile':
        t.media.add(label);
        break;
      case 'contentPart':
        t.ink.add(label);
        break;
      case 'blipFill':
        picture(e);
        break;
      case 'bodyPr':
        textBody(e);
        break;
      case 'pPr':
        paragraph(e);
        break;
      case 'rPr':
        run(e);
        break;
      case 'fld':
        t.fields.add(label);
        break;
      case 'tbl':
        table(e);
        break;
      case 'tc':
        if (Number(e.attrs.gridSpan) >= 2 || Number(e.attrs.rowSpan) >= 2) t.merge.add(label);
        break;
      case 'tcPr':
        if (e.attrs.vert !== undefined && e.attrs.vert !== 'horz') t.vertical.add(label);
        break;
      case 'm':
        if (e.ns === A14) t.equations.add(label);
        break;
      case 'hlinkClick':
      case 'hlinkHover':
        if (e.attrs.action) t.actions.add(actionKind(e.attrs.action));
        break;
      case 'transition':
        transition = true;
        break;
      case 'timing':
        timing = true;
        break;
      case 'pic':
        // a picture's crop-to-shape preset is not counted, but custom geometry is
        if (kid(kid(e, 'spPr'), 'custGeom')) {
          t.custGeom.add(label);
          t.geometry.add('custGeom');
        }
        break;
      case 'cTn':
        if (e.attrs.presetClass !== undefined) effects++;
        break;
      case 'overrideClrMapping':
        t.colorMap.add(label);
        break;
      case 'gradFill':
        if (parent && FILL_HOLDERS.has(parent.name)) t.gradientFill.add(label);
        break;
      case 'pattFill':
        if (parent && FILL_HOLDERS.has(parent.name)) t.patternFill.add(label);
        break;
      case 'effectLst':
      case 'effectDag':
        if (parent && EFFECT_HOLDERS.has(parent.name) && !shadows.has(e) && (e.name === 'effectDag' || e.children.length > 0))
          t.effects.add(label);
        break;
      default:
        if (ZOOM.has(e.name)) t.zoom.add(label);
        else if (ANIMATES.has(e.name)) animates = true;
    }
    const nests = e.name === 'grpSp';
    if (nests) groupDepth++;
    for (const k of e.children) if (!(k.name === 'Fallback' && k.ns === 'mc')) visit(k, e);
    if (nests) groupDepth--;
  };

  visit(root, undefined);
  if (!isOn(root.attrs.show)) t.hidden.add(label);
  if (transition) t.transitions.add(label);
  if (effects > 0) t.animations.add(label, effects);
  else if (timing && animates) t.animations.add(label);
}
