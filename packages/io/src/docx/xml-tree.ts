import { XMLParser } from 'fast-xml-parser';

// ---------------------------------------------------------------------------
// XML tree utilities. Parsed with preserveOrder: true — required because a
// document body interleaves <w:p> and <w:tbl> children in document order;
// fast-xml-parser's default mode groups same-tag siblings into separate
// arrays and loses that interleaving entirely. The tradeoff is a more
// verbose node shape: each element is `{ tagName: [children], ':@': attrs }`.
// ---------------------------------------------------------------------------

export type XmlNode = Record<string, unknown> & { ':@'?: Record<string, string> };

const xmlParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
});

export function parseXml(xml: string): XmlNode[] {
  return xmlParser.parse(xml) as XmlNode[];
}

/** The single non-attribute key of a preserveOrder node, e.g. "w:p". */
export function tagOf(node: XmlNode): string | undefined {
  return Object.keys(node).find((k) => k !== ':@');
}

export function childrenOf(node: XmlNode): XmlNode[] {
  const tag = tagOf(node);
  if (!tag) return [];
  const value = node[tag];
  return Array.isArray(value) ? (value as XmlNode[]) : [];
}

export function attrsOf(node: XmlNode): Record<string, string> {
  return node[':@'] ?? {};
}

export function findAll(nodes: XmlNode[], tag: string): XmlNode[] {
  return nodes.filter((n) => tagOf(n) === tag);
}

export function find(nodes: XmlNode[], tag: string): XmlNode | undefined {
  return nodes.find((n) => tagOf(n) === tag);
}


/** Pre-order (document-order) walk of every element; `path` holds the ancestor tags, outermost first. */
export function walk(nodes: XmlNode[], visit: (node: XmlNode, tag: string, path: readonly string[]) => void, path: string[] = []): void {
  for (const node of nodes) {
    const tag = tagOf(node);
    if (!tag || tag === '#text') continue;
    visit(node, tag, path);
    path.push(tag);
    walk(childrenOf(node), visit, path);
    path.pop();
  }
}
