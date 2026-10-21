import { readFile } from 'node:fs/promises';
import path from 'node:path';
import type { Page } from '@playwright/test';

export const FIXTURE_DIR = path.resolve(__dirname, '..', 'fixtures');
export const fixture = (name: string): string => path.join(FIXTURE_DIR, name);

/** Collect uncaught page exceptions (must stay empty in every file-I/O test). */
export function collectPageErrors(page: Page): string[] {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  return errors;
}

/**
 * Record `window.alert` text (the demos report open/download failures with an alert)
 * and accept it. Returns the live list of messages.
 */
export function collectAlerts(page: Page): string[] {
  const messages: string[] = [];
  page.on('dialog', (d) => {
    messages.push(d.message());
    void d.accept();
  });
  return messages;
}

/** Click the toolbar's Open button and feed the file chooser (the real user path). */
export async function openViaChooser(
  page: Page,
  buttonName: string,
  file: string | { name: string; mimeType: string; buffer: Buffer }
): Promise<void> {
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.getByRole('button', { name: buttonName, exact: true }).click(),
  ]);
  await chooser.setFiles(file);
}

export interface Downloaded {
  fileName: string;
  bytes: Buffer;
}

/** Run `trigger` and return the bytes of the file the page downloaded. */
export async function captureDownload(
  page: Page,
  trigger: () => Promise<void>
): Promise<Downloaded> {
  const [download] = await Promise.all([page.waitForEvent('download'), trigger()]);
  const filePath = await download.path();
  return { fileName: download.suggestedFilename(), bytes: await readFile(filePath) };
}

export function asUpload(d: Downloaded, mimeType: string) {
  return { name: d.fileName, mimeType, buffer: d.bytes };
}
