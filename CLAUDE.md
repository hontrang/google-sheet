# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

TypeScript codebase that powers a Google Sheet for Vietnamese stock market (VN-INDEX) analysis. The TypeScript is compiled/copied into Google Apps Script (GAS) projects and deployed via `clasp`. Most identifiers (functions, variables, sheet names) are in Vietnamese.

## The three deployable "projects"

The same shared utilities feed three separate GAS projects, each with its own `code.ts` + `trigger.ts` and its own external clasp folder (paths set in `.env`):

- `src/projectMain` — **Dashboard** (`build:main`). Per-stock detail scraping, news, financial reports, charts.
- `src/projectDw` — **Data Warehouse** (`build:dw`). VN-INDEX index data and aggregate market metrics.
- `src/projectTest` — **Test** project run inside GAS (`build:test`), bundles `tests/gas/*.ts` (QUnit).

Only one project lives in `dist/` at a time — whatever you last built. `npm run deploy` pushes whatever is currently in `dist/`.

## Build & deploy

```sh
npm run build:main          # build Dashboard into dist/
npm run build:dw            # build Data Warehouse into dist/
npm run build:test          # build Test project into dist/
npm run deploy              # cd dist && clasp push && clasp open  (pushes current dist/)
npm run build:main:deploy   # build + deploy in one step (also :dw, :test variants)
npm run lint                # eslint .
npm run prettier            # format all .ts
```

### How the build pipeline works (important)

GAS has no module system, so the build flattens everything into one namespace:

1. `clean` — wipe `dist/` and `build/`.
2. `build:assets` / `build:utility` — copy `src/assets/lib/*.ts` and `src/utility/*.ts` into `dist/`.
3. `cpMain` / `cpDW` / `cpTest` — copy the chosen project's `code.ts` + `trigger.ts` into `dist/`.
4. `removeImport` (`src/scripts/remove-import.ts`) — **comments out every `import` statement** in `dist/*.ts` via gulp regex. This is why path aliases and cross-file imports work in source but vanish in the deployed code; all helper classes must resolve as globals at runtime in GAS.
5. `prettier`.
6. `build:clasp:*` (`src/scripts/build-clasp-*.ts`) — copy clasp metadata (`.clasp.json`, `appsscript.json`, etc.) from the external folder named by the corresponding `*Clasp` env var into `dist/`.

Build scripts run with `bun`. Path aliases (`tsconfig.json`): `@utils/* → src/utility/*`, `@src/* → src/*`.

## Online vs offline (HTTP abstraction)

`IHttp` (`src/types/generic.ts`) has two implementations — keep them interchangeable:

- `src/utility/HttpHelper.ts` — production, uses GAS `UrlFetchApp`.
- `src/offline/AxiosHelper.ts` — local/Node, uses `axios` (and SSI token auth). Pairs with `src/offline/ExcelHelper.ts` for reading/writing local `.xlsx` instead of a live Sheet.

The `mode` env var (`offline`/online) selects the path. `src/utility/SheetHelper.ts` wraps all sheet I/O; sheet names live in `SheetHelper.sheetName`.

## Tests

Playwright is used purely as a test runner (no browser). Projects defined in `playwright.config.ts`:

```sh
npm run test                              # all *.test.ts
npx playwright test http.test.ts         # single file
npx playwright test --project offline    # offline crawl (tests/crawl.offline.ts)
npx playwright test --project tradingeconomics
```

Copy the provided `.env` into the project root before testing (it holds clasp folder paths and SSI credentials).

## CI (.github/workflows)

- `ci.yml` — on `v1.*.*` tag push: runs `http.test.ts`, creates a GitHub Release.
- `tradingeconomics.yml` — daily cron: runs the `tradingeconomics` project, publishes `docs/` to GitHub Pages, posts a Discord webhook.

## Conventions

- ESLint extends typescript-eslint strict + stylistic + type-checked. Naming convention enforces `camelCase` for object literal properties, with an explicit allow-list for HTTP header keys (`Accept`, `Cookie`, `User-Agent`, `__RequestVerificationToken`, etc.) in `eslint.config.js`.
- Data is fetched/scraped from several Vietnamese finance sources (vps.com.vn, cafef.vn, vndirect, simplize.vn, vietstock); cheerio parses HTML. Response shapes are typed as `IResponse*` interfaces in `src/types/generic.ts`.
- GAS triggers are declared in each project's `trigger.ts` (`mainCreateTriggers` / `mainDeleteTrigger`); they only take effect once deployed.