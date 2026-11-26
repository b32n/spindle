import { walk, type XmlElement } from '../xlsx/xml';

/**
 * Families the editor can show without the file carrying them: the web-safe set every browser maps to an installed font
 * (Arial, Times New Roman, Courier New, Georgia, Verdana, Tahoma, Trebuchet MS and their generic aliases) and the
 * Office defaults that have metric-compatible open substitutes (Calibri -> Carlito, Cambria -> Caladea, plus the Liberation
 * and Noto families those substitutes ship with). The list is deliberately short: a font outside it is reported as
 * `pptx.fonts.unavailable`, which only means "the deck may not look the same here", never that it is lost.
 * Names are compared case-insensitively. Theme references (`+mj-lt`, `+mn-ea`) are not font names and never count.
 */
export const AVAILABLE_FONTS: ReadonlySet<string> = new Set(
  [
    'Arial', 'Helvetica', 'Times New Roman', 'Times', 'Courier New', 'Courier', 'Georgia', 'Verdana', 'Tahoma',
    'Trebuchet MS', 'Calibri', 'Calibri Light', 'Cambria', 'Carlito', 'Caladea',
    'Liberation Sans', 'Liberation Serif', 'Liberation Mono', 'Noto Sans', 'Noto Serif', 'DejaVu Sans',
    'sans-serif', 'serif', 'monospace', 'system-ui',
  ].map((f) => f.toLowerCase())
);

export const isAvailableFont = (name: string): boolean => AVAILABLE_FONTS.has(name.trim().toLowerCase());

/** Typeface attributes of `latin`, `ea` and `cs` elements anywhere under `root` (rich-text defaults, runs, theme font schemes). */
export function collectFonts(root: XmlElement | undefined, into: Set<string>): void {
  walk(root, (e) => {
    if (e.name !== 'latin' && e.name !== 'ea' && e.name !== 'cs') return;
    const face = (e.attrs.typeface ?? '').trim();
    if (face !== '' && !face.startsWith('+')) into.add(face);
  });
}
