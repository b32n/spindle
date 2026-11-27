import type { SpreadsheetImportResult, XlsxSharedString, XlsxStyles, XlsxTheme, XlsxWorkbookInfo } from '@b32nio/spindle-shared/export/xlsx';
import type { ColorContext } from './colors';
import { parseSharedStrings } from './shared-strings';
import { parseStyles } from './styles';
import { parseTheme } from './theme';
import type { XlsxPackage } from './package';
import { readWorkbook } from './workbook';

// ---------------------------------------------------------------------------
// The own XLSX package reader: the foundation layer (S2a) and, on top of it, the sheet reader (`sheet.ts`, S2b) that `importXlsx`
// (`../reader.ts`) is built from. Internal API, not exported from the package entry points.
//
//   const pkg = await XlsxPackage.open(bytes, { limits });  // X4 safe zip: only ever throws ImportError
//   const parts = await readWorkbookParts(pkg);              // workbook, theme, styles, shared strings; never throws after open,
//                                                           //   except ImportError('corrupt') when there is no workbook part
//   readSheet(await pkg.text(ref.part), ref, ctx)            // a sheet part (S2b), with `ref.rels` for its relationships
//
// Streaming-ready (S20): every part reader is a pure function from a parsed part to plain, structured-cloneable data, with no
// DOM and no shared state, so the whole reader can run in a worker. Workbook-level parts are small and use the bounded element
// tree. `sheetData` is not read here: `readSheet` reads it as a forward-only event stream (`sax.ts`, no tree) from `pkg.text()`, which holds the part as one string. Shared strings can be
// large, so `forEachSharedString` pushes items to a callback instead of requiring the list.
// ---------------------------------------------------------------------------

export { XlsxPackage, decodePart, type OpenXlsxOptions } from './package';
export { parseContentTypes, contentTypeOf, partsWithContentType, type ContentTypes } from './content-types';
export { parseRelationships, relsOfKind, resolveRelationshipTarget, relsPart } from './relationships';
export { findWorkbookPart, isBinaryWorkbook, readWorkbook } from './workbook';
export { parseTheme } from './theme';
export { PRESET_COLORS } from './preset-colors';
export { parseStyles, numberFormatCode } from './styles';
export { readFont } from './font';
export { parseSharedStrings, forEachSharedString, readSharedString, decodeStringEscapes } from './shared-strings';
export { resolveColor, applyTint, parseRgb, readColor, DEFAULT_INDEXED_COLORS, DEFAULT_THEME_COLORS, type ColorContext } from './colors';

export interface XlsxWorkbookParts {
  workbook: XlsxWorkbookInfo;
  theme: XlsxTheme;
  styles: XlsxStyles;
  /** Indexed as cells refer to them; empty when the workbook has no shared-strings part. */
  sharedStrings: XlsxSharedString[];
  /** What a colour resolves against (theme colours and the workbook's palette), for the sheet reader's own colour elements (tab colour, CF). */
  colors: ColorContext;
  /** Non-fatal problems found up to this call (a copy of the package's warnings, in order; later reads do not change it). */
  warnings: string[];
}

/** Read the workbook-level parts of an opened package. */
export async function readWorkbookParts(pkg: XlsxPackage): Promise<XlsxWorkbookParts> {
  const workbook = await readWorkbook(pkg);
  const theme = parseTheme(workbook.parts.theme === undefined ? undefined : await pkg.xml(workbook.parts.theme), pkg.warnings);
  const styles = parseStyles(workbook.parts.styles === undefined ? undefined : await pkg.xml(workbook.parts.styles), theme.colors);
  const colors: ColorContext = { themeColors: theme.colors, indexedColors: styles.indexedColors };
  const stringsRoot = workbook.parts.sharedStrings === undefined ? undefined : await pkg.xml(workbook.parts.sharedStrings);
  return { workbook, theme, styles, sharedStrings: parseSharedStrings(stringsRoot, colors), colors, warnings: [...pkg.warnings] };
}

/** The own reader's output in the shape of the io result (`SpreadsheetImportResult`'s optional S2 fields); `importXlsx` adds it to its result. */
export function toImportExtras(parts: XlsxWorkbookParts): Required<Pick<SpreadsheetImportResult, 'workbook' | 'styles' | 'sharedStrings' | 'theme'>> {
  return { workbook: parts.workbook, styles: parts.styles, sharedStrings: parts.sharedStrings, theme: parts.theme };
}
