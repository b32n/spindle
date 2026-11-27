import { XMLParser } from 'fast-xml-parser';

// ---------------------------------------------------------------------------
// The one XML parser configuration for importers (fast-xml-parser v5), with its safety limits in one place.
//
// - Entities are never expanded: `processEntities: false`, and a DOCTYPE (with its internal subset, where entities
//   are declared) is cut out before parsing, so an entity bomb ("billion laughs") or an external entity is inert. This
//   is stricter than the parser's own capped expansion (`processEntities: { maxTotalExpansions, maxExpandedLength, ... }`),
//   which is not needed because nothing here has a use for custom entities. The five predefined entities and numeric
//   references are decoded by the caller in one pass (`decodeEntities` in xlsx/xml.ts).
// - Nesting is capped by `maxNestedTags` (verified against fast-xml-parser 5.12: the option exists, throws
//   "Maximum nested tags exceeded" and applies to the tag depth of the document).
// - Parsing never throws: `parseXmlTree` answers `null` for anything it cannot or will not read.
// ---------------------------------------------------------------------------

/** XML limits, shared by every importer (SD16: depth 256). */
export const XML_LIMITS = { maxDepth: 256 };

/**
 * A parser with the safe settings; `preserveOrder` keeps element order, which OOXML depends on. With `decodePredefined`
 * the parser itself decodes the five predefined entities (as the DOCX reader always did); that is only safe on text whose
 * DOCTYPE is removed first (`stripDoctype`), because then no entity can be declared, so none can expand.
 */
export function createSafeXmlParser(maxDepth: number = XML_LIMITS.maxDepth, decodePredefined = false): XMLParser {
  return new XMLParser({
    preserveOrder: true,
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    trimValues: false,
    parseTagValue: false,
    parseAttributeValue: false,
    processEntities: decodePredefined,
    // set here so a change of the parser's default cannot change the safety behaviour
    maxNestedTags: maxDepth,
  });
}

/** Remove a DOCTYPE declaration (with its internal subset). Returns `null` when it cannot be delimited. */
export function stripDoctype(xml: string): string | null {
  const start = xml.indexOf('<!DOCTYPE');
  if (start < 0) return xml;
  const open = xml.indexOf('[', start);
  const close = xml.indexOf('>', start);
  if (close < 0) return null;
  if (open < 0 || open > close) return xml.slice(0, start) + xml.slice(close + 1);
  const end = /\]\s*>/.exec(xml.slice(open));
  return end ? xml.slice(0, start) + xml.slice(open + end.index + end[0].length) : null;
}

const defaultParser = createSafeXmlParser();

/** The order-preserving raw tree of `xml`, or `null` when it is not XML, too deeply nested or has an undelimited DOCTYPE. Never throws. */
export function parseXmlTree(xml: string, parser: XMLParser = defaultParser): unknown[] | null {
  try {
    const clean = stripDoctype(xml);
    if (clean === null) return null;
    return parser.parse(clean) as unknown[];
  } catch {
    return null;
  }
}
