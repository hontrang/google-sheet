# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

**Before working on this repo, read [architecture.md](architecture.md).** It has the full component diagrams, the annotated directory tree (§2), the exact `dist/` → GAS push mechanics (§2.4), the build/deploy sequence, and — critically — the actual Google Sheet IDs/names bound to each GAS project and how their tabs map to `SheetHelper.sheetName`. Re-read it whenever a task touches the build pipeline, `dist/`, a specific project's sheet structure, or the online/offline abstraction, since this file only summarizes.

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
2. `build:assets` / `build:utility` — copy `src/assets/lib/*.ts` into `dist/zAssets/` (flattened, `-f`) and `src/utility/*.ts` into `dist/utility/` (path kept, `-a`). **`dist/` is not fully flat.**
3. `cpMain` / `cpDW` / `cpTest` — copy the chosen project's `code.ts` + `trigger.ts` to the `dist/` root; `cpTest` also copies `tests/gas/*.ts` into `dist/gas/`.
4. `removeImport` (`src/scripts/remove-import.ts`) — comments out `import` statements via gulp regex, but its glob is `dist/*.ts`, so it only touches **root-level** files. `dist/utility/*.ts` (and `dist/gas/*.ts`) keep their live imports. That is harmless: ts2gas comments imports out again at push time (see below). Use `dist/**/*.ts` if you ever need the subfolders handled here.
5. `prettier`.
6. `build:clasp:*` (`src/scripts/build-clasp-*.ts`) — copy clasp metadata from the external folder named by the corresponding `*Clasp` env var into `dist/`: `.clasp.json`, `.claspignore`, `appsscript.json`, `package.json`, plus `.clasprc.json` + `creds.json` for `main` only.

Build scripts run with `bun`. Path aliases (`tsconfig.json`): `@utils/* → src/utility/*`, `@src/* → src/*` — source-only; see below for why they never reach GAS.

### How `dist/` gets pushed to GAS (important)

`npm run deploy` = `cd dist && clasp push && clasp open`, so **clasp treats `dist/` as the project root** and reads `.clasp.json` / `.claspignore` / `appsscript.json` from there, not from the repo root. Verify what a push would send with `cd dist && clasp status` (read-only, no network).

- **Pushed**: `appsscript.json`, `code.ts`, `trigger.ts`, `utility/*.ts`, `zAssets/*.ts`. **Ignored**: `.clasp.json`, `.claspignore`, `.clasprc.json`, `creds.json`, `package.json`, `.DS_Store`. clasp only accepts `.ts`/`.js`/`.gs` (→ `SERVER_JS`), `.html`, and `appsscript.json` as the sole JSON, so the credential files can sit in `dist/` without being uploaded.
- **Subfolders become filenames**: `dist/utility/SheetHelper.ts` lands as the GAS file `utility/SheetHelper.gs`. The editor shows `/` as pseudo-folders, but everything shares **one global scope** at runtime. File order is load order (no `filePushOrder` is set, so: root files first, then subfolders alphabetically).
- **clasp transpiles, not tsc**: each `.ts` goes through **ts2gas** (bundled in clasp 2.4.2, TypeScript 4.9.5, default `target: ES3`, `module: None`). clasp only overrides those options from a `tsconfig.json` **inside the project root**, and `dist/` has none — so the repo's `tsconfig.json` has **zero effect on the deployed output**; it is for the IDE, eslint and `tsc` only. ts2gas also comments out every `import` and emits `var exports/module` shims, which is why classes end up as hoisted global `var`s.
- **`dist/package.json` (`{}`) must exist**: ts2gas reads `package.json` from the cwd on load, and the cwd is `dist/`. Delete it and `clasp push` dies with `ENOENT: ... open 'package.json'`.
- **`.clasp.json` picks the target project** — that copy step is the entire "which project am I deploying" mechanism. Three distinct scriptIds (main / dw / test); the folders live outside the repo and are not committed, so a fresh clone needs `.env` plus those folders before it can build.
- **A push replaces the whole project**: clasp calls `script.projects.updateContent` with the full file list, so files that exist on GAS but not in `dist/` are **deleted**. Pushing the wrong `dist/` wipes the target project rather than merging into it — always build immediately before deploying, or use `build:main:deploy` / `build:dw:deploy` / `build:test:deploy`.
- Pushing does **not** register triggers; `mainCreateTriggers` in `trigger.ts` must be run once inside GAS.

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
- `tradingeconomics.yml` — daily cron: runs the `tradingeconomics` project, publishes `publish/` to GitHub Pages (branch `gh-pages`), posts a Discord webhook. `publish/` holds generated output, not documentation — [tradingeconomics.ts](tests/tradingeconomics.ts) overwrites `publish/index.html` on every run. Full write-up incl. known pitfalls: [docs/publish-workflow.md](docs/publish-workflow.md).

## Conventions

- ESLint extends typescript-eslint strict + stylistic + type-checked. Naming convention enforces `camelCase` for object literal properties, with an explicit allow-list for HTTP header keys (`Accept`, `Cookie`, `User-Agent`, `__RequestVerificationToken`, etc.) in `eslint.config.js`.
- Data is fetched/scraped from several Vietnamese finance sources (vps.com.vn, cafef.vn, vndirect, simplize.vn, vietstock); cheerio parses HTML. Response shapes are typed as `IResponse*` interfaces in `src/types/generic.ts`.
- GAS triggers are declared in each project's `trigger.ts` (`mainCreateTriggers` / `mainDeleteTrigger`); they only take effect once deployed.