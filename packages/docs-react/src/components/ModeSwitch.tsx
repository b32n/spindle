/**
 * ModeSwitch - Editing / Suggesting / Viewing.
 */

import { memo } from 'react';

/** What typing does: `editing` changes the document, `suggesting` records tracked changes, `viewing` cannot change it. */
export type EditorMode = 'editing' | 'suggesting' | 'viewing';

export interface ModeSwitchProps {
  mode: EditorMode;
  onModeChange: (mode: EditorMode) => void;
}

const MODES: Array<{ id: EditorMode; label: string; hint: string }> = [
  { id: 'editing', label: 'Editing', hint: 'Edit the document directly' },
  { id: 'suggesting', label: 'Suggesting', hint: 'Your edits become tracked changes (Track Changes)' },
  { id: 'viewing', label: 'Viewing', hint: 'Read the document without changing it' },
];

export const ModeSwitch = memo(function ModeSwitch({ mode, onModeChange }: ModeSwitchProps) {
  return (
    <div
      role="group"
      aria-label="Editing mode"
      data-testid="mode-switch"
      style={{ display: 'inline-flex', border: '1px solid #dadce0', borderRadius: 6, overflow: 'hidden', background: '#fff', font: '12px/16px Arial, sans-serif' }}
    >
      {MODES.map((m) => (
        <button
          key={m.id}
          type="button"
          title={m.hint}
          aria-pressed={mode === m.id}
          onClick={() => onModeChange(m.id)}
          style={{
            border: 'none',
            borderRight: m.id === 'viewing' ? 'none' : '1px solid #dadce0',
            padding: '3px 10px',
            cursor: 'pointer',
            background: mode === m.id ? (m.id === 'suggesting' ? '#fef7e0' : '#e8f0fe') : '#fff',
            color: mode === m.id ? (m.id === 'suggesting' ? '#b06000' : '#1967d2') : '#3c4043',
            fontWeight: mode === m.id ? 600 : 400,
          }}
        >
          {m.label}
        </button>
      ))}
    </div>
  );
});
