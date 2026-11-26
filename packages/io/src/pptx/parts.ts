/**
 * Part names the detector knows: package plumbing and properties, the parts a PPTX importer will read (slides, layouts,
 * masters, themes), and the parts one of the loss ids covers (charts, diagrams, media, embeddings, comments, notes, ink,
 * macros, signatures). A part matching none of these is dropped without a specific report: `pptx.package.unaccountedParts`.
 */
const KNOWN_PARTS: readonly RegExp[] = [
  /^\[Content_Types\]\.xml$/i,
  /(?:^|\/)_rels\/[^/]*\.rels$/,
  /^docProps\//,
  /^ppt\/(?:presentation|presProps|viewProps|tableStyles)\.xml$/,
  /^ppt\/(?:slides|slideLayouts|slideMasters)\/[^/]+\.xml$/,
  /^ppt\/theme\//,
  /^ppt\/printerSettings\//,
  /^ppt\/(?:notesSlides|notesMasters|handoutMasters|charts|diagrams|media|embeddings|comments|ink|activeX)\//,
  /^ppt\/(?:commentAuthors|authors)\.xml$/,
  /^ppt\/vbaProject(?:Signature)?\.bin$/,
  /^_xmlsignatures\//,
];

export function isKnownPart(name: string): boolean {
  return KNOWN_PARTS.some((re) => re.test(name));
}

export const SLIDE = /^ppt\/slides\/[^/]+\.xml$/;
export const LAYOUT = /^ppt\/slideLayouts\/[^/]+\.xml$/;
export const MASTER = /^ppt\/slideMasters\/[^/]+\.xml$/;
export const NOTES_SLIDE = /^ppt\/notesSlides\/[^/]+\.xml$/;
export const COMMENTS = /^ppt\/comments\/[^/]+\.xml$/;

/** The last path segment: `ppt/slides/slide3.xml` -> `slide3.xml`. */
export const baseName = (part: string): string => part.slice(part.lastIndexOf('/') + 1);
