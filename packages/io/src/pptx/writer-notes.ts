import type { PptxColorMap, PptxTextBody } from '@b32nio/spindle-shared/export/pptx';
import { txBodyXml } from './writer-body';
import type { WriteContext } from './writer-context';
import { COLOR_MAP_ROLES } from './writer-layouts';
import { attrs, ROOT_NS, TREE_HEADER, XML_DECLARATION } from './writer-xml';

// Speaker notes (plan K13): the notes master (the notes page's slide image and body placeholders) and one notes slide per slide that
// has notes. The inverse of `readNotes`. The text carries its own effective style, so the notes master states no text styles.

/** The notes page is the portrait size `p:notesSz` states: the slide image above, the notes body below it. */
const slideImage = (ph: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph${ph}/></p:nvPr></p:nvSpPr>`;

const SLIDE_IMAGE_FRAME = '<a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>';
const BODY_FRAME = '<a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';

const bodyShape = (body: string, frame: string): string =>
  '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="1"/></p:nvPr></p:nvSpPr>' +
  `<p:spPr>${frame}</p:spPr><p:txBody>${body}</p:txBody></p:sp>`;

/** `ppt/notesMasters/notesMaster1.xml`: the slide image and the notes body, the colour map of the deck. */
export function notesMasterXml(colorMap: PptxColorMap): string {
  const map = attrs(Object.fromEntries(COLOR_MAP_ROLES.map((r) => [r, colorMap[r]])));
  return (
    `${XML_DECLARATION}<p:notesMaster ${ROOT_NS}><p:cSld><p:spTree>${TREE_HEADER}${slideImage(' type="sldImg" idx="2"')}<p:spPr>${SLIDE_IMAGE_FRAME}</p:spPr></p:sp>` +
    `${bodyShape('<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p>', BODY_FRAME)}</p:spTree></p:cSld><p:clrMap${map}/></p:notesMaster>`
  );
}

/** `ppt/notesSlides/notesSlideN.xml`: the slide image placeholder and the notes body, placed by the notes master. */
export function notesSlideXml(notes: PptxTextBody, ctx: WriteContext): string {
  return (
    `${XML_DECLARATION}<p:notes ${ROOT_NS}><p:cSld><p:spTree>${TREE_HEADER}${slideImage(' type="sldImg"')}<p:spPr/></p:sp>` +
    `${bodyShape(txBodyXml(notes, ctx), '')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`
  );
}
