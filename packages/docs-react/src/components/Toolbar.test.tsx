import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { DocumentImpl } from '@b32nio/spindle-docs-core';
import { DocumentProvider } from '../context/DocumentContext';
import { Toolbar, type ToolbarProps } from './Toolbar';

function render(props: ToolbarProps): string {
  return renderToStaticMarkup(
    <DocumentProvider document={new DocumentImpl('d', 'T')}>
      <Toolbar {...props} />
    </DocumentProvider>
  );
}

describe('docs Toolbar file actions', () => {
  // ResponsiveToolbar uses useLayoutEffect, which React warns about under
  // server rendering; static markup is all these tests need.
  beforeAll(() => {
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('shows no file controls by default', () => {
    const html = render({});
    expect(html).not.toContain('docs-open-file-input');
    expect(html).not.toContain('Download');
    expect(html).not.toContain('Open file');
  });

  it('shows Open with a local .docx file input when onOpenFile is given', () => {
    const html = render({ onOpenFile: () => {} });
    expect(html).toContain('Open file');
    expect(html).toContain('type="file"');
    expect(html).toContain('accept=".docx"');
  });

  it('shows a direct Download button when only one format is available', () => {
    const html = render({ onDownload: () => {} });
    expect(html).toContain('Download as Microsoft Word (.docx)');
    expect(html).not.toContain('type="file"');
  });
});
