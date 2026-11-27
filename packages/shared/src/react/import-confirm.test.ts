import { createElement as h } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import type { ImportLossItem, ImportLossReport } from '../export/loss';

// The dialog portals into document.body; render it in place for static markup.
jest.mock('react-dom', () => ({ ...jest.requireActual('react-dom'), createPortal: (node: unknown) => node }));
import { ImportConfirmDialog } from './import-confirm';

(globalThis as { document?: unknown }).document = { body: {} }; // only the portal target; effects do not run in static markup

const item = (id: string, state?: ImportLossItem['state']): ImportLossItem => ({
  id, label: id, count: 2, effect: 'because', ...(state && { state }),
});
const render = (items: ImportLossItem[]) =>
  renderToStaticMarkup(h(ImportConfirmDialog, { fileName: 'f.pptx', report: { format: 'pptx', items } as ImportLossReport, onConfirm() {}, onCancel() {} }));

describe('ImportConfirmDialog', () => {
  it('renders a report without states as one "will be lost" list (unchanged)', () => {
    const html = render([item('a'), item('b')]);
    expect(html).toContain('and the following will be lost:');
    expect(html).toContain('data-testid="import-loss-list"');
    expect(html).not.toContain('import-loss-kept-list');
    expect(html.match(/<li/g)).toHaveLength(2);
    expect(html).toContain('2 × a — because');
  });

  it('groups kept and lost items separately once any item has a state', () => {
    const html = render([item('chart', 'kept'), item('macro'), item('fmt', 'approximated')]);
    expect(html).toContain('Will be lost:');
    expect(html).toContain('Kept, not editable (saved back unchanged):');
    expect(html).toContain('Imported approximately:');
    const lost = html.slice(html.indexOf('data-testid="import-loss-list"'), html.indexOf('Imported approximately'));
    expect(lost).toContain('macro');
    expect(lost).not.toContain('chart');
    const kept = html.slice(html.indexOf('data-testid="import-loss-kept-list"'));
    expect(kept).toContain('2 × chart — because');
    expect(kept).not.toContain('macro');
  });

  it('omits the empty groups', () => {
    const html = render([item('chart', 'kept')]);
    expect(html).not.toContain('Will be lost:');
    expect(html).not.toContain('import-loss-list"');
    expect(html).not.toContain('approximated');
    expect(html).toContain('import-loss-kept-list');
  });
});
