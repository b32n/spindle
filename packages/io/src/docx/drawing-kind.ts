import { attrsOf, findDeep, type XmlNode } from './xml-tree';

export type DrawingKind = 'picture' | 'chart' | 'diagram' | 'shape' | 'embeddedObject' | 'unknown';

/** Classifies a <w:drawing> by its <a:graphicData uri="...">, which says what the drawing really is. A drawing with no graphicData that still carries a blip is treated as a picture. */
export function drawingKind(drawingNode: XmlNode): DrawingKind {
  const graphicData = findDeep(drawingNode, 'a:graphicData');
  const uri = (graphicData ? attrsOf(graphicData)['@_uri'] : undefined)?.toLowerCase() ?? '';
  if (uri.endsWith('/drawingml/2006/picture')) return 'picture';
  if (uri.includes('/chart')) return 'chart'; // .../drawingml/2006/chart, .../drawing/2014/chartex
  if (uri.endsWith('/drawingml/2006/diagram')) return 'diagram';
  if (uri.includes('wordprocessingshape') || uri.includes('wordprocessinggroup') || uri.includes('wordprocessingcanvas') || uri.endsWith('/drawingml/2006/lockedcanvas')) {
    return 'shape';
  }
  if (/\/ole(\/|$)/.test(uri) || (graphicData && findDeep(graphicData, 'o:OLEObject'))) return 'embeddedObject';
  if (!uri && findDeep(drawingNode, 'a:blip')) return 'picture';
  return 'unknown';
}

