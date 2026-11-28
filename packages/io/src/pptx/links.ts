import type { PptxLink } from '@b32nio/spindle-shared/export/pptx';
import { kid, type XmlElement } from '../xlsx/xml';
import type { Rels } from './fill';

// Hyperlinks (plan K10): the scheme allow-list of the editor's links (http, https, mailto and tel; the S10 rule). Linked targets are
// never fetched and `ppaction://` and `file:` targets are kept inert.

/** The schemes a link may have (the writer checks against the same list). */
export const LINK_SCHEMES = new Set(['http', 'https', 'mailto', 'tel']);
const SCHEME = /^([a-z][a-z0-9+.-]*):/i;
// control characters and whitespace inside the scheme are how `java\tscript:` slips past a naive check
// eslint-disable-next-line no-control-regex
const STRIP = /[\u0000- \u007f-\u009f]/g;

export interface LinkContext {
  rels: Rels;
  warn: (message: string) => void;
}

/** The scheme of an address (lower case, whitespace and control characters ignored), `undefined` for a relative one. */
export const schemeOf = (target: string): string | undefined => SCHEME.exec(target.replace(STRIP, ''))?.[1].toLowerCase();

/** `a:hlinkClick` as a link; `undefined` when there is none or its target is dropped (with a warning that names only the scheme). */
export function readLink(hlink: XmlElement | undefined, ctx: LinkContext): PptxLink | undefined {
  if (!hlink) return undefined;
  const tooltip = (hlink.attrs.tooltip ?? '').trim();
  const extra = tooltip !== '' ? { tooltip } : {};
  const action = hlink.attrs.action ?? '';
  const rel = ctx.rels.get(hlink.rattrs?.id ?? '');
  const slidePart = rel && !rel.external && rel.type === 'slide' ? rel.target : undefined;
  if (/^ppaction:\/\//i.test(action.replace(STRIP, ''))) return { kind: 'action', target: action, ...(slidePart ? { slidePart } : {}), ...extra };
  if (!rel) return undefined;
  if (!rel.external) return slidePart ? { kind: 'action', target: action || 'ppaction://hlinksldjump', slidePart, ...extra } : undefined;
  const target = rel.raw.trim();
  const scheme = schemeOf(target);
  if (scheme !== undefined && LINK_SCHEMES.has(scheme)) return { kind: 'url', href: target, ...extra };
  if (scheme === 'file') {
    ctx.warn('A link to a file is kept as an inert action; it is never opened.');
    return { kind: 'action', target, ...extra };
  }
  ctx.warn(scheme === undefined ? 'A hyperlink with a relative address was dropped.' : `A hyperlink with a "${scheme.slice(0, 20)}:" address was dropped; only http, https, mailto and tel links are kept.`);
  return undefined;
}

/** The `hlinkClick` of the `cNvPr` of a shape, picture or frame. */
export const shapeLink = (cNvPr: XmlElement | undefined, ctx: LinkContext): PptxLink | undefined => readLink(kid(cNvPr, 'hlinkClick'), ctx);
