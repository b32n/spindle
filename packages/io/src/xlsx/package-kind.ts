/** Package-level facts the loss detector reads without opening any XML: what kind of file this is, and which parts it knows. */

const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** An OLE compound file: a legacy `.xls`, or an OOXML workbook wrapped in an encryption container. */
export function isOleFile(bytes: Uint8Array): boolean {
  return bytes.length >= OLE_MAGIC.length && OLE_MAGIC.every((b, i) => bytes[i] === b);
}

/** True when `bytes` holds `text` as UTF-16LE (how a compound file stores its stream names). */
function hasUtf16(bytes: Uint8Array, text: string): boolean {
  const first = text.charCodeAt(0);
  const last = bytes.length - text.length * 2;
  for (let i = 0; i <= last; i++) {
    if (bytes[i] !== first || bytes[i + 1] !== 0) continue;
    let k = 1;
    while (k < text.length && bytes[i + 2 * k] === text.charCodeAt(k) && bytes[i + 2 * k + 1] === 0) k++;
    if (k === text.length) return true;
  }
  return false;
}

/** A compound file that holds an encrypted package (password to open) rather than a legacy BIFF workbook. */
export function isEncryptedOle(bytes: Uint8Array): boolean {
  return hasUtf16(bytes, 'EncryptedPackage') || hasUtf16(bytes, 'EncryptionInfo');
}

/**
 * Part names the importer reads, reports under a loss id, or can safely ignore (package plumbing and document
 * properties). A part matching none of these is dropped without a specific report: `xlsx.package.unaccountedParts`.
 */
const KNOWN_PARTS: readonly RegExp[] = [
  /^\[Content_Types\]\.xml$/i,
  /(?:^|\/)_rels\/[^/]*\.rels$/,
  /^docProps\//,
  /^xl\/(?:workbook|sharedStrings|styles|calcChain|metadata)\.xml$/,
  /^xl\/worksheets\/[^/]+\.xml$/,
  /^xl\/theme\//,
  /^xl\/printerSettings\//,
  // each of these is covered by a loss id (charts, images, objects, comments, tables, pivots, slicers, links, macros, data)
  /^xl\/comments[^/]*\.xml$/,
  /^xl\/(?:chartsheets|drawings|charts|media|comments|threadedComments|persons|tables|pivotTables|pivotCache|slicers|slicerCaches|externalLinks|ctrlProps|activeX|embeddings|queryTables|richData|model)\//,
  /^xl\/(?:connections|cellimages)\.xml$/,
  /^xl\/vbaProject(?:Signature)?\.bin$/,
];

export function isKnownPart(name: string): boolean {
  return KNOWN_PARTS.some((re) => re.test(name));
}
