import { diffValues, isObj } from '../fidelity-core/diff';
import type { DiffReport, Json, GroupedCounts } from '../fidelity-core/types';
import { resolveStyles } from './normalize';

/** Plain text of a block (inline text, link text, table cells and the blocks inside them, content controls and what they wrap), used to align blocks. */
function blockText(block: Json): string {
  if (!isObj(block)) return '';
  const inline = (items: Json | undefined): string =>
    Array.isArray(items)
      ? items.map((i) => (isObj(i) && typeof i.text === 'string' ? i.text : isObj(i) && Array.isArray(i.content) ? inline(i.content) : isObj(i) && i.type ? `[${String(i.type)}]` : '')).join('')
      : '';
  if (Array.isArray(block.blocks)) return block.blocks.map(blockText).join('');
  if (Array.isArray(block.rows)) {
    return block.rows
      .map((row) =>
        isObj(row) && Array.isArray(row.cells)
          ? row.cells.map((c) => (isObj(c) ? (Array.isArray(c.blocks) ? c.blocks.map(blockText).join('\n') : inline(c.content)) : '')).join('\t')
          : ''
      )
      .join('\n');
  }
  return inline(block.content);
}

const signature = (block: Json): string => `${isObj(block) ? String(block.type) : '?'}|${blockText(block)}`;

function countModel(model: Json): GroupedCounts {
  let sections = 0;
  let textChars = 0;
  const blocks: { [type: string]: number } = {};
  if (isObj(model) && Array.isArray(model.sections)) {
    sections = model.sections.length;
    for (const section of model.sections) {
      if (!isObj(section) || !Array.isArray(section.blocks)) continue;
      for (const block of section.blocks) {
        const type = isObj(block) ? String(block.type) : 'unknown';
        blocks[type] = (blocks[type] ?? 0) + 1;
        textChars += blockText(block).length;
      }
    }
  }
  return { sections, blocks: Object.fromEntries(Object.entries(blocks).sort(([x], [y]) => x.localeCompare(y))), textChars };
}

/**
 * Structural diff of two normalized docs-core models. Styles are resolved first (so pool keys do not
 * matter) and section block lists are aligned by content, so one dropped paragraph is one entry, not
 * a cascade of index shifts.
 */
export function diffModels(before: Json, after: Json): DiffReport {
  return diffValues(before, after, {
    counts: countModel,
    prepare: resolveStyles,
    signatures: (path, b, a) => (path.endsWith('.blocks') ? { before: b.map(signature), after: a.map(signature) } : null),
  });
}
