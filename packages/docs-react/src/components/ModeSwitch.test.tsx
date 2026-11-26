import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { ModeSwitch, type EditorMode } from './ModeSwitch';

const render = (mode: EditorMode) => renderToStaticMarkup(<ModeSwitch mode={mode} onModeChange={() => {}} />);

describe('ModeSwitch', () => {
  it('offers Editing, Suggesting and Viewing, and marks the one that is on', () => {
    for (const [mode, label] of [['editing', 'Editing'], ['suggesting', 'Suggesting'], ['viewing', 'Viewing']] as const) {
      const html = render(mode);
      expect(html).toContain('aria-label="Editing mode"');
      for (const other of ['Editing', 'Suggesting', 'Viewing']) expect(html).toContain(`>${other}<`);
      expect(html).toMatch(new RegExp(`aria-pressed="true"[^>]*>${label}<`));
      expect((html.match(/aria-pressed="true"/g) ?? []).length).toBe(1);
    }
  });

  it('says in a hint what each mode does, and names Track Changes for Suggesting', () => {
    const html = render('editing');
    expect(html).toContain('Edit the document directly');
    expect(html).toContain('Track Changes');
    expect(html).toContain('Read the document without changing it');
  });
});
