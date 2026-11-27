import { kids, type XmlElement } from '../xml';

/** `[Content_Types].xml`: a default type per extension and an override per part (ECMA-376 Part 2, 10.1.2). */
export interface ContentTypes {
  /** Lower-cased extension (no dot) to content type. */
  defaults: ReadonlyMap<string, string>;
  /** Part name (no leading slash, lower-cased) to content type. */
  overrides: ReadonlyMap<string, string>;
}

export const EMPTY_CONTENT_TYPES: ContentTypes = { defaults: new Map(), overrides: new Map() };

/** Parse the content-types root. Entries without the needed attributes are skipped; a missing root gives no types. */
export function parseContentTypes(root: XmlElement | undefined): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const d of kids(root, 'Default')) {
    const ext = d.attrs.Extension;
    const type = d.attrs.ContentType;
    if (ext !== undefined && type !== undefined && !defaults.has(ext.toLowerCase())) defaults.set(ext.toLowerCase(), type);
  }
  for (const o of kids(root, 'Override')) {
    const name = o.attrs.PartName;
    const type = o.attrs.ContentType;
    if (name === undefined || type === undefined) continue;
    const key = name.replace(/^\/+/, '').toLowerCase();
    if (!overrides.has(key)) overrides.set(key, type);
  }
  return { defaults, overrides };
}

/** The content type of a part: its override, else the default of its extension, else `undefined`. */
export function contentTypeOf(types: ContentTypes, part: string): string | undefined {
  const override = types.overrides.get(part.toLowerCase());
  if (override !== undefined) return override;
  const dot = part.lastIndexOf('.');
  return dot < 0 ? undefined : types.defaults.get(part.slice(dot + 1).toLowerCase());
}

/** The parts (as written in the overrides, lower-cased) that carry a content type satisfying `test`. */
export function partsWithContentType(types: ContentTypes, test: (type: string) => boolean): string[] {
  const found: string[] = [];
  for (const [part, type] of types.overrides) if (test(type)) found.push(part);
  return found;
}
