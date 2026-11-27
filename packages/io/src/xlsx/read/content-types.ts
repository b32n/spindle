import { scanXml } from './sax';
import type { ListOptions } from './relationships';

/** `[Content_Types].xml`: a default type per extension and an override per part (ECMA-376 Part 2, 10.1.2). */
export interface ContentTypes {
  /** Lower-cased extension (no dot) to content type. */
  defaults: ReadonlyMap<string, string>;
  /** Part name (no leading slash, lower-cased) to content type. */
  overrides: ReadonlyMap<string, string>;
}

export const EMPTY_CONTENT_TYPES: ContentTypes = { defaults: new Map(), overrides: new Map() };

/**
 * Parse the text of `[Content_Types].xml`, read as events (memory is the entries kept, at most `options.max`). Entries without the
 * needed attributes are skipped; no text gives no types. Damaged markup ends the read there and an entry over the limit ends it too
 * (`options.problem`).
 */
export function parseContentTypes(xml: string | undefined, options: ListOptions = {}): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  if (xml === undefined) return { defaults, overrides };
  const max = options.max ?? Infinity;
  let depth = 0;
  let over = false;
  const wellFormed = scanXml(xml, {
    open(name, attrs) {
      depth++;
      if (depth !== 2) return;
      const type = attrs.ContentType;
      if (type === undefined) return;
      if (name === 'Default') {
        const ext = attrs.Extension;
        if (ext === undefined || defaults.has(ext.toLowerCase())) return;
        if (defaults.size + overrides.size >= max) over = true;
        else defaults.set(ext.toLowerCase(), type);
      } else if (name === 'Override') {
        const part = attrs.PartName;
        if (part === undefined) return;
        const key = part.replace(/^\/+/, '').toLowerCase();
        if (overrides.has(key)) return;
        if (defaults.size + overrides.size >= max) over = true;
        else overrides.set(key, type);
      }
    },
    close() {
      depth--;
    },
    halted: () => over,
  });
  if (over) options.problem?.('over');
  else if (!wellFormed) options.problem?.('damaged');
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
