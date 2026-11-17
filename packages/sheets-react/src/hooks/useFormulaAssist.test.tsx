import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { useFormulaAssist, type FormulaAssist } from './useFormulaAssist';

function capture(value: string, caret: number, enabled = true): FormulaAssist {
  let result!: FormulaAssist;
  function Probe() {
    result = useFormulaAssist({ value, caret, enabled, onAccept: () => {} });
    return null;
  }
  renderToStaticMarkup(<Probe />);
  return result;
}

describe('useFormulaAssist (derived state)', () => {
  it('offers autocomplete for a partial function name with the cursor on the first item', () => {
    const assist = capture('=SU', 3);
    expect(assist.mode).toBe('autocomplete');
    expect(assist.suggestions.length).toBeGreaterThan(0);
    expect(assist.highlightedIndex).toBe(0);
  });

  it('stays hidden when disabled or when the value is not a formula', () => {
    expect(capture('=SU', 3, false).mode).toBe('hidden');
    expect(capture('hello', 5).mode).toBe('hidden');
  });

  it('shows the signature while inside a call', () => {
    expect(capture('=SUM(1,', 7).mode).toBe('signature');
  });
});
