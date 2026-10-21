import type { Locator, Page } from '@playwright/test';

/**
 * Toolbar button that falls back to the "More tools" overflow menu when the
 * responsive toolbar has collapsed it (the demo toolbars overflow at any
 * viewport width since Open/Download were added). Only `click()` is adapted;
 * visibility assertions still target the inline button.
 */
export function toolbarButton(page: Page, name: string | RegExp): Locator {
  const exact = typeof name === 'string';
  const button = page.getByRole('button', { name, exact });
  const click = button.click.bind(button);
  button.click = async (options) => {
    if ((await button.count()) > 0) return click(options);
    await page.getByRole('button', { name: 'More tools', exact: true }).click();
    const menu = page.getByRole('menu');
    await menu
      .getByRole('button', { name })
      .or(menu.getByRole('menuitem', { name }))
      .first()
      .click(options);
  };
  return button;
}

/** Dismiss the "More tools" overflow menu if a toolbar action left it open. */
export async function closeToolbarMenu(page: Page): Promise<void> {
  const menu = page.getByRole('menu');
  if (await menu.isVisible()) {
    await page.mouse.click(2, 2);
    await menu.waitFor({ state: 'hidden' });
  }
}
