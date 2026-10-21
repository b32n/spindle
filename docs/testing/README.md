# Testing guide

Spindle has two layers of automated tests. Use the cheapest one that can catch the bug.

| Layer | Runner | Lives in | Use it for |
|---|---|---|---|
| Unit | Jest (per package) | `packages/*/src/**/*.test.ts` | The formula engine, selection/commands logic, collab bindings, and **file round-trips (CSV / XLSX / DOCX import & export)**. Fast, no browser. |
| E2E | Playwright | `e2e/` | Anything that needs a real browser: pointer gestures, focus, keyboard shortcuts, canvas/DOM interaction, clipboard, dialogs, present mode, and two-client collaboration. |

> **Rule:** the formula engine and file round-trips stay in fast unit tests. Do not
> re-test them through the browser; E2E only proves the UI is wired to them.

> **Rule:** every bug fix ships with a regression test. If the bug is observable in the
> browser (almost every UI bug), that is an E2E test that fails before the fix and passes after.
> Commit the failing test first, then the fix, so the history shows the reproduction.

The full list of scenarios (with stable IDs and status) is in [`scenarios.md`](./scenarios.md).

## Running unit tests

```bash
npm ci
npm run build          # cross-package types resolve through each package's dist/
npm test               # all workspaces
npm test -w @b32nio/spindle-sheets-core -- formula   # one package, name filter
```

## Running E2E tests

Playwright starts the demo apps itself (`playwright.config.ts` → `webServer`):
sheets `:5175`, docs `:5176`, slides `:5177` (Vite, aliased to package `src/`, so no build is needed)
and the reference collab relay on `ws://127.0.0.1:1234`.

```bash
npx playwright install --with-deps chromium firefox webkit   # once
npm run test:e2e                      # everything, all browsers
npm run test:e2e:smoke                # only @smoke scenarios
npx playwright test --project=chromium e2e/sheets/selection.spec.ts
npx playwright test -g "S-SEL-07"     # by scenario ID
npm run test:e2e:ui                   # Playwright UI mode (time-travel, locator picker)
npx playwright test --repeat-each=3 --project=chromium   # stability check before merging
```

Servers are reused locally if already running (`npm run dev` in an example), and always started fresh in CI.

### Debugging

- **UI mode** (`npm run test:e2e:ui`): watch a test step by step, pick locators, re-run on save.
- **Headed + inspector:** `npx playwright test <file> --headed --debug`.
- **Traces:** failures keep a trace under `e2e/.results/<test>/trace.zip`; open with
  `npx playwright show-trace <zip>`. In CI download the `e2e-failure-artifacts-N` artifact
  (traces, screenshots, videos) or the merged `playwright-html-report`.
- **HTML report:** `npx playwright show-report e2e/.report`.

## How the editors are driven

- **Sheets** is a `<canvas>`; cells have no DOM. The page object (`e2e/support/sheets.ts`) converts
  `(row, col)` to pixels (header 100×20, cells 100×20). State is read from
  `data-selection` / `data-active-cell` on the canvas (`data-testid="sheet-canvas"`), the name box
  (`data-testid="name-box"`) and the model handle `window.__spindle.workbook`.
- **Docs** edits a hidden ProseMirror instance (`.ProseMirror`, off-screen) and paints pages as DOM
  (`.page`, `.page-content`). Click rendered text, then type; assert on the hidden editor's DOM.
- **Slides** renders elements as DOM (`[data-element-id]`, `[data-slide-stage]`, `[data-handle]`);
  the model handle is `window.__spindle.deck`.

The demos expose `window.__spindle` purely as a test hook (no network, no telemetry).

## Writing stable tests (no sleeps)

- Use role / test-id locators and **web-first assertions** (`await expect(locator).toHaveText(...)`,
  `expect.poll(...)`). Never use `page.waitForTimeout`.
- Wait on an observable state, not time. Examples already in the support files:
  `gotoSlide()` waits for the stage to show the slide; docs `clickText()`/`typeText()`/`press()`
  wait for the ProseMirror selection to be mirrored before the next key (ProseMirror applies
  selection asynchronously, so keys sent immediately act on a stale caret).
- Collaboration tests use one room per test (`room = ...${testInfo.testId}`), a fresh
  browser context per peer, wait for the "· connected" banner, and use a *witness* peer before
  adding a late joiner. Compare text with the remote-cursor widgets stripped (docs `headingText`).
- Chromium-only features (real clipboard read/write) use `test.skip(browserName !== 'chromium', reason)`.

## Adding a scenario

1. Pick the area file in `e2e/<editor>/` (or add one) and an unused ID: `S-` sheets, `D-` docs, `SL-` slides,
   then an area code and number, e.g. `S-FRZ-03`.
2. Write `test('S-FRZ-03 freeze survives undo', async ...)`. Add `@smoke` to the title only for
   the handful of fast, high-value happy paths.
3. If the product is broken, write the test anyway, mark it `test.fixme('ID title', ...)`, and put a
   `// Bug: ...` comment directly above it describing the bug plus the issue URL. File the issue first.
4. Regenerate the map: `npm run docs:scenarios` (CI-friendly check: `npm run docs:scenarios:check`).
   Add intentional gaps to `GAPS` in `scripts/testing/scenarios.mjs`.
5. Run it repeatedly: `npx playwright test -g "S-FRZ-03" --repeat-each=10 --workers=4`.

## Fixing a bug (regression-test workflow)

1. Reproduce in a new E2E test (or unit test for pure logic). Commit it failing.
2. Fix the root cause. Commit.
3. The test (and, where sensible, a unit test) must pass; remove any `fixme`.

## CI

`.github/workflows/e2e.yml` runs on every pull request and push to `master`: Node 22, `npm ci`, build,
cached Playwright browsers, 4 shards, blob reports merged into an HTML report, and
traces/screenshots/videos uploaded on failure. It also builds the three demos for production and runs the
`@smoke` subset on Chromium against `vite preview` (`npm run build:demos` then `npm run test:e2e:preview`
locally). The aggregate job **`e2e-result`** is the single check to require (see below); it fails if any
shard or the production smoke failed, or if `docs/testing/scenarios.md` is stale.

**Flaky tests fail the run.** CI keeps `retries: 2` so a failure still produces traces, but
`failOnFlakyTests` makes a test that only passed on a retry fail its shard; the flaky tests are listed in the
job summary of the *Merge reports* job. Fix the test (or product race), do not retry around it.

`.github/workflows/e2e-nightly.yml` (nightly + manual `workflow_dispatch`) runs everything on all three
browsers with `--retries=0 --repeat-each=5` in 6 shards and lists every test that failed in any repeat.
Reproduce locally with `npx playwright test -g "ID" --repeat-each=10 --workers=4`.

### Making `e2e-result` required

Repository → **Settings → Branches → Branch protection rules** (or **Rules → Rulesets**) for `master`:
enable *Require status checks to pass before merging*, search for **`e2e-result`** (and the existing
`Packages (build · type-check · test · lint)` / `Website (type-check · build)` checks), select them, save.
The check name only appears in the search box after the workflow has run once on the repository.

## Local git hooks

`npm install` runs the `prepare` script, which installs `simple-git-hooks`
(skipped in CI, outside a git checkout, or with `SKIP_GIT_HOOKS=1`; it never fails an install).

- **pre-commit** (`lint-staged`): ESLint `--fix` on staged `packages/*/src`, then per affected package
  `tsc --noEmit` and `jest --findRelatedTests`; Prettier + e2e type-check for staged `e2e/` files.
- **pre-push**: nothing by default. `SPINDLE_PREPUSH_E2E=1 git push` runs the `@smoke` E2E subset first.
- Bypass once with `git commit --no-verify` (CI still runs everything).
