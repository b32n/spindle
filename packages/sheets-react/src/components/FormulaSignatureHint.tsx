import React, { memo } from 'react';
import type { FunctionDoc } from '@b32nio/spindle-sheets-core';

interface FormulaSignatureHintProps {
  doc: FunctionDoc;
  /** 0-based index of the argument currently being entered. */
  activeArg: number;
  /** Absolute position (within the editor's positioned container). */
  top: number;
  left: number;
}

/** Tooltip showing a function's signature with the active argument bold. */
export const FormulaSignatureHint = memo(function FormulaSignatureHint({
  doc,
  activeArg,
  top,
  left,
}: FormulaSignatureHintProps) {
  return (
    <div
      style={{
        position: 'absolute',
        top,
        left,
        zIndex: 1100,
        pointerEvents: 'none',
        maxWidth: '440px',
        padding: '6px 10px',
        background: '#ffffff',
        border: '1px solid rgba(15, 23, 42, 0.12)',
        borderRadius: '8px',
        boxShadow: '0 8px 24px rgba(15, 23, 42, 0.16)',
        fontFamily: '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif',
        fontSize: '13px',
      }}
    >
      <div style={{ color: '#64748b' }}>
        <span style={{ fontWeight: 600, color: '#1e293b' }}>{doc.name}</span>(
        {doc.args.map((arg, index) => {
          // The trailing variadic argument stays active for every repeat of it.
          const isActive =
            index === activeArg || (arg.variadic === true && activeArg >= index);
          let label = arg.name;
          if (arg.variadic) label += ', ...';
          if (arg.optional) label = `[${label}]`;
          return (
            <React.Fragment key={index}>
              {index > 0 && ', '}
              <span
                style={{
                  fontWeight: isActive ? 700 : 400,
                  color: isActive ? '#1e293b' : '#64748b',
                }}
              >
                {label}
              </span>
            </React.Fragment>
          );
        })}
        )
      </div>
      <div style={{ marginTop: '4px', color: '#94a3b8', fontSize: '11px' }}>
        {doc.description}
      </div>
    </div>
  );
});
