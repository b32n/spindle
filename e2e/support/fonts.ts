import type { Page } from '@playwright/test';

const isLocal = (url: string) =>
  /^(https?:\/\/(127\.0\.0\.1|localhost)[:/]|data:|blob:|about:)/.test(url);

/**
 * Serves stylesheet requests to `cssHosts` with an empty stylesheet (200) or, with `fail`, aborts them,
 * aborts every other external request, and returns the list of external URLs the page asked for.
 * Register before `goto`.
 */
export async function trackExternalRequests(
  page: Page,
  opts: { cssHosts?: string[]; fail?: boolean } = {}
): Promise<string[]> {
  const seen: string[] = [];
  page.on('request', (req) => {
    if (!isLocal(req.url())) seen.push(req.url());
  });
  await page.route(
    (url) => !isLocal(url.href),
    (route) => {
      const host = new URL(route.request().url()).host;
      if (!opts.fail && opts.cssHosts?.includes(host)) {
        return route.fulfill({
          status: 200,
          contentType: 'text/css',
          body: '/* test stylesheet */',
        });
      }
      return route.abort();
    }
  );
  return seen;
}

export const FONT_LINKS = 'link[data-spindle-fonts]';
