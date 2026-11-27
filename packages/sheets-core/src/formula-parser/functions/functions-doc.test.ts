import * as fs from 'fs';
import * as path from 'path';
import { getAllFunctionDocs } from './catalog';

// The published function list (documentation/sheets/core/functions.md) has to list every function of the catalog, in its category.
const doc = fs.readFileSync(path.resolve(__dirname, '../../../../../documentation/sheets/core/functions.md'), 'utf8');

describe('published function list', () => {
  it('states the catalog size', () => {
    expect(doc).toContain(`ships **${getAllFunctionDocs().length} built-in functions**`);
  });

  it('lists every catalog function once, under its category, with its signature', () => {
    const sections = new Map<string, string>();
    for (const part of doc.split(/^## /m).slice(1)) sections.set(part.split('\n')[0], part);
    const titles: Record<string, string> = {
      math: 'Math and aggregation',
      logical: 'Logical',
      text: 'Text',
      lookup: 'Lookup',
      stats: 'Statistical',
      date: 'Date and time',
      information: 'Information',
      financial: 'Financial',
      reference: 'Reference',
      array: 'Array (dynamic arrays)',
    };
    const docs = getAllFunctionDocs();
    for (const [category, title] of Object.entries(titles)) {
      const names = docs.filter((d) => d.category === category).map((d) => d.name);
      const section = [...sections].find(([heading]) => heading.startsWith(`${title} (`));
      expect(section?.[0]).toBe(`${title} (${names.length})`);
      const listed = [...(section as [string, string])[1].matchAll(/^\| `([A-Z0-9.]+)\(/gm)].map((m) => m[1]);
      expect(listed.slice().sort()).toEqual(names.slice().sort());
    }
  });

  it('names the evidence behind every function', () => {
    for (const line of doc.split('\n').filter((l) => /^\| `[A-Z0-9.]+\(/.test(l))) {
      expect(line).toMatch(/\| (unit tests|\d+ oracle rows, [^|]+) \|$/);
    }
  });
});
