import { createElement as h } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { ResponsiveToolbar } from './ResponsiveToolbar';

describe('ResponsiveToolbar', () => {
  beforeAll(() => {
    // useLayoutEffect warns under server rendering; static markup is all we need.
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterAll(() => {
    jest.restoreAllMocks();
  });

  it('renders every child inline before any measurement (no overflow trigger)', () => {
    const html = renderToStaticMarkup(
      h(ResponsiveToolbar, null, h('button', null, 'One'), h('button', null, 'Two'))
    );
    expect(html).toContain('One');
    expect(html).toContain('Two');
    expect(html).not.toContain('More tools');
  });
});
