// Excel number-format engine: `formatValue(value, code)` renders a cell value the way Excel would display it.
// Codes are compiled once and cached. See documentation/sheets/number-formats.md for the supported grammar and the
// places where this differs from Excel or LibreOffice on purpose.

import { generalDigits } from './digits';
import { parseFormat, type CompiledFormat } from './parse';
import { OVERFLOW, renderNumber, renderText, selectSection, type FormatOptions, type FormatResult } from './render';

export { dateToSerial, serialToDate } from './dates';
export type { CalendarDate, DateSystem } from './dates';
export type { CompiledFormat, Section, Token } from './parse';
export type { FormatOptions, FormatResult } from './render';

const CACHE_LIMIT = 500;
const cache = new Map<string, CompiledFormat>();

/** Parse (and cache) a format code. Never throws; see `ok` / `error` on the result. */
export function compileFormat(code: string): CompiledFormat {
  let compiled = cache.get(code);
  if (!compiled) {
    if (cache.size >= CACHE_LIMIT) cache.clear();
    compiled = parseFormat(code);
    cache.set(code, compiled);
  }
  return compiled;
}

export function isValidFormatCode(code: string): boolean {
  return compileFormat(code).ok;
}

/** True when the code's first section shows dates or times (so a number in it means a date serial). */
export function isDateTimeFormatCode(code: string): boolean {
  const compiled = compileFormat(code);
  return compiled.ok && compiled.sections[0].kind === 'date';
}

const ERROR_TEXT = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|ERROR!)/;

/** The text of a number in General format (also what text-only sections show for a number). */
export function formatGeneral(value: number, full = false): string {
  if (!Number.isFinite(value)) return '#NUM!';
  return (value < 0 ? '-' : '') + generalDigits(Math.abs(value), value < 0, full);
}

/**
 * Render `value` with an Excel format code. Numbers use the sections; text uses the text section; booleans and
 * errors pass through unformatted; null is empty. An invalid code shows the value in General (`invalid: true`).
 */
export function formatValue(value: unknown, code: string, options: FormatOptions = {}): FormatResult {
  if (value === null || value === undefined) return { text: '' };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' };
  const compiled = compileFormat(code);
  const invalid = compiled.ok ? {} : { invalid: true };
  if (typeof value === 'string') return { ...textResult(compiled, value), ...invalid };
  if (typeof value !== 'number' || !Number.isFinite(value)) return { text: typeof value === 'number' ? '#NUM!' : String(value) };
  const sections = compiled.sections;
  const { index, minus } = selectSection(sections, value);
  if (index < 0) return { text: formatGeneral(value, options.fullGeneral), ...invalid };
  const section = sections[index];
  const rendered = renderNumber(section, value, minus, options);
  const result: FormatResult = { text: rendered ? rendered.text : OVERFLOW, ...invalid };
  if (section.color) result.color = section.color;
  if (rendered?.fill) result.fill = rendered.fill;
  return result;
}

function textResult(compiled: CompiledFormat, text: string): FormatResult {
  if (ERROR_TEXT.test(text)) return { text };
  const sections = compiled.sections;
  const section = sections.length === 4 ? sections[3] : sections[sections.length - 1].hasText ? sections[sections.length - 1] : undefined;
  if (!section) return { text };
  const rendered = renderText(section, text);
  const result: FormatResult = { text: rendered.text };
  if (section.color) result.color = section.color;
  if (rendered.fill) result.fill = rendered.fill;
  return result;
}

/** `formatValue(...).text` */
export function formatDisplay(value: unknown, code: string, options: FormatOptions = {}): string {
  return formatValue(value, code, options).text;
}
