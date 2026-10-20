import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { Toolbar } from './Toolbar';

describe('sheets Toolbar file actions', () => {
  // ResponsiveToolbar uses useLayoutEffect, which React warns about under
  // server rendering; static markup is all these tests need.
  beforeAll(() => {
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('shows no file controls by default', () => {
    const html = renderToStaticMarkup(<Toolbar />);
    expect(html).not.toContain('sheets-open-file-input');
    expect(html).not.toContain('Download');
    expect(html).not.toContain('Open file');
  });

  it('shows Open (with a local .xlsx/.csv file input) when onOpenFile is given', () => {
    const html = renderToStaticMarkup(<Toolbar onOpenFile={() => {}} />);
    expect(html).toContain('Open file');
    expect(html).toContain('type="file"');
    expect(html).toContain('accept=".xlsx,.csv"');
  });

  it('shows the Download control when onDownload is given', () => {
    const html = renderToStaticMarkup(<Toolbar onDownload={() => {}} />);
    expect(html).toContain('Download');
    expect(html).not.toContain('type="file"');
  });
});
