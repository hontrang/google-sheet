# Báo cáo lỗi code — 2026-10-31

Ảnh chụp tình trạng lỗi TypeScript (`tsc`) và ESLint của repo, kèm cách lọc bỏ lỗi từ thư viện/output không liên quan. Số liệu đo ngày 2026-10-31, sau khi đã sửa `tsconfig.json` cho TypeScript 6 (`types` khai báo rõ, bỏ `baseUrl`).

## 0. Cập nhật 2026-10-31 — đã sửa hết lỗi trong code của repo

| | `tsc` TS 5.5.4 | `tsc` TS 6.0.3 | ESLint |
| --- | ---: | ---: | ---: |
| Lỗi thật trong code của repo | **0** (từ 18) | **0** (từ 18) | **0** (từ 12) |
| Còn lại — toàn bộ thuộc thư viện/output (mục 2) | 7.195 | 7.173 | 418 (406 thư viện + 12 lỗi parse `dist/`, `src/generator/`, `eslint.config.js`) |

Đã kiểm tra lại sau khi sửa: `build:dw`, `build:test`, `build:main` chạy đúng (đúng scriptId, bước comment `import` vẫn chạy), `playwright test --list` vẫn thấy 20 test. Muốn các con số tổng về 0, áp dụng mục 2 (chỉ đổi cấu hình, không sửa thư viện).

## 1. Tóm tắt (lúc lập báo cáo)

| | `tsc` (TS 5.5.4 của repo) | `tsc` (TS 6.0.3, VS Code) | ESLint (`npm run lint`) |
| --- | ---: | ---: | ---: |
| Tổng hiện tại | 7.213 | 7.191 | 430 |
| — trong đó `src/assets/lib/` (thư viện copy vào repo) | 7.195 | 7.173 | 406 |
| — trong đó file không thuộc tsconfig (`dist/`, `src/generator/`, `eslint.config.js`) | — | — | 12 (lỗi parse) |
| **Lỗi thật trong code của repo** | **18** | **18** | **12** |

- `projectMain`, `projectDw`, `utility`, `types`: **0 lỗi `tsc`**.
- 99,7% số lỗi là nhiễu từ `src/assets/lib/` (`Cheerio.ts`, `luxon.ts` — ~980 KB code bundle của bên thứ ba) và từ `dist/` (output build).
- 30 lỗi thật tập trung ở `tests/`, `src/projectTest/`, `src/scripts/` và `src/offline/ExcelHelper.ts` — không có lỗi nào trong code chạy production trên GAS.

## 2. Cách ignore lỗi thư viện / output

Hai thay đổi cấu hình, **không sửa code thư viện**. Đã chạy thử bằng config tạm: kết quả đúng như cột "Lỗi thật" ở trên (18 / 18 / 12).

### 2.1 `tsconfig.json` — loại `src/assets/lib` khỏi type-check

```jsonc
"exclude": [
  "node_modules",
  "**/*.spec.ts",
  "src/assets/lib"
]
```

`src/assets/lib/*.ts` chỉ được build copy nguyên vào `dist/zAssets/` để push lên GAS, không file nào trong `src/` hay `tests/` import nó, nên loại khỏi tsconfig không ảnh hưởng gì tới build/deploy (clasp cũng không đọc `tsconfig.json` — xem [CLAUDE.md](../CLAUDE.md)).

### 2.2 `eslint.config.js` — thêm khối `ignores` toàn cục

Đặt làm phần tử đầu tiên trong `tseslint.config(...)`:

```js
{
  ignores: ['dist/**', 'src/assets/lib/**', 'src/generator/**', 'eslint.config.js']
},
```

| Pattern | Lý do |
| --- | --- |
| `dist/**` | Output build (copy của `src/`), lint ở đây là lint trùng; lại không thuộc tsconfig nên ESLint báo lỗi parse |
| `src/assets/lib/**` | Thư viện bên thứ ba (239 lỗi `no-var`, 54 `naming-convention`, …) |
| `src/generator/**` | Thư mục webpack cũ có `node_modules` riêng, file `.js` không thuộc tsconfig |
| `eslint.config.js` | Không thuộc tsconfig → lỗi parse với `parserOptions.project` |

### 2.3 Lưu ý

- Sau khi đổi, chạy **TypeScript: Restart TS Server** trong VS Code. Mở trực tiếp file trong `src/assets/lib` vẫn có thể thấy lỗi (VS Code dùng "inferred project" cho file ngoài tsconfig) — bình thường, không ảnh hưởng `tsc`/lint.
- Nếu sau này cập nhật Cheerio/luxon, chỉ cần thay file trong `src/assets/lib`, không phải lo lỗi lint.

## 3. Danh sách 18 lỗi TypeScript (đã sửa hết)

| File:dòng | Mã | Lỗi | Nguyên nhân / hướng sửa |
| --- | --- | --- | --- |
| ~~`src/projectTest/code.ts:23,24`~~ | TS2304 | Không tìm thấy `TestSheet`, `TestRequest` | ✅ Đã sửa — `declare function TestSheet(): void` / `TestRequest` (hàm nằm trong `tests/gas/*.ts`, là module vì có `import`) |
| ~~`src/projectTest/code.ts:34,57,67,78`~~ | TS2304 | Không tìm thấy `QUnitGS2` | ✅ Đã sửa — `declare const QUnitGS2: { QUnit; init; getHtml; getResultsFromServer }`; thêm `declare global { interface Config { title: string } }` cho `QUnit.config.title` (lỗi lộ ra khi `QUnit` hết là `any`) |
| ~~`src/scripts/build-clasp-{dw,main,test}.ts:7`, `src/scripts/clean.ts:7`~~ | TS1470 | `import.meta` không được phép khi output CommonJS | ✅ Đã sửa — thay `fileURLToPath(import.meta.url)` + `dirname` bằng `__dirname` (bun hỗ trợ sẵn); đã chạy `build:dw`, `build:test`, `build:main` — đúng scriptId, `clean` vẫn xoá `dist/` |
| ~~`src/scripts/remove-import.ts:6`~~ | TS2741 | `ReadWriteStream` thiếu `compose` | ✅ Đã sửa — kiểu trả về đổi thành `NodeJS.ReadWriteStream` (đúng kiểu `gulp.src().pipe()` trả về), bỏ import `Stream` |
| ~~`tests/crawl.offline.ts:46,47`~~ | TS18046 | `response.data` có kiểu `unknown` | ✅ Đã sửa — ép `response.data as { message: string; data: IResponseSsi[] }` |
| ~~`tests/gas/TestRequest.ts:10`~~ | TS2552 | Không tìm thấy `URLFetchRequestOptions` | ✅ Đã sửa — dùng `GoogleAppsScript.URL_Fetch.URLFetchRequestOptions` |
| ~~`tests/gas/TestSheetHelper.ts:37,45,53,106`~~ | TS2345 | `number[]` truyền vào chỗ cần `string[]` | ✅ Đã sửa — `assert.deepEqual<unknown[]>(…)`; kiểu `string[]` của `layDuLieuTrongHang` chưa đổi (đổi sẽ lan sang `ISheetSpread` và mọi chỗ gọi) |

## 4. Danh sách 12 lỗi ESLint (đã sửa hết)

| File:dòng | Rule | Ghi chú |
| --- | --- | --- |
| ~~`src/offline/ExcelHelper.ts:44`~~ | `no-non-null-assertion` | ✅ Đã sửa — bỏ `!`, ném `Error('Không tìm thấy sheet …')` nếu sheet không tồn tại (trước đây trả `undefined` rồi lỗi ở chỗ khác) |
| ~~`src/offline/ExcelHelper.ts:87,95,100,122,127`~~ | `no-explicit-any` | ✅ Đã sửa — `any` → `CellValue` (exceljs) cho `ghiDuLieuVaoCot`, `ghiDuLieuVaoDayTheoVung`, `ghiDuLieuVaoO`; `unknown[]`/`unknown[][]` cho `ghiDuLieuVaoDay`, `ghiDuLieuVaoDayTheoTen` |
| ~~`src/projectTest/code.ts:34`~~ | `no-var` | ✅ Đã sửa — `const QUnit`. TS coi file là module (`module: NodeNext`) nên không đụng `const QUnit` global của `@types/qunit`; output ts2gas (ES3) vẫn là `var QUnit = …`, đã so sánh giống hệt bản trên `main` |
| ~~`src/projectTest/code.ts:55`~~ | `no-unused-vars` | ✅ Đã sửa — `doGet` do GAS gọi; thêm `/* eslint-disable @typescript-eslint/no-unused-vars */` đầu file (theo quy ước của `trigger.ts`, `TestSheetHelper.ts`) |
| ~~`src/projectTest/code.ts:77`~~ | `no-unused-vars` | ✅ Đã sửa — `getResultsFromServer` do client QUnit gọi; như trên |
| ~~`tests/gas/TestRequest.ts:1`~~ | `no-unused-vars` | ✅ Đã sửa — xoá import `IResponseVietStock` thừa |
| ~~`tests/gas/TestRequest.ts:5`~~ | `no-unused-vars` | ✅ Đã sửa — `TestRequest` dùng qua global scope GAS; thêm `eslint-disable` đầu file |
| ~~`src/utility/DateHelper.ts:4`~~ | `no-extraneous-class` | ✅ Đã sửa — giữ class static (GAS cần một tên global `DateHelper`), tắt rule bằng `eslint-disable-next-line` kèm lý do, như `AxiosHelper.ts` |

## 5. Đề xuất thứ tự xử lý

1. **Áp dụng mục 2** (tsconfig `exclude` + ESLint `ignores`) — 5 phút, đưa `npm run lint` và `tsc` từ hàng nghìn lỗi về đúng 30 lỗi thật, để lỗi mới không bị chìm.
2. **Lỗi dễ, không rủi ro**: import thừa (`TestRequest.ts:1`), `URLFetchRequestOptions`, generic axios trong `crawl.offline.ts`, `eslint-disable` cho các hàm GAS gọi (`doGet`, `getResultsFromServer`, `TestRequest`).
3. **Khai báo type cho GAS-only globals**: `QUnitGS2`, `TestSheet`, `TestRequest` trong một file `.d.ts`.
4. **`src/scripts/`**: quyết định `.mts` hay bỏ `import.meta`; kiểm tra lại `npm run build:*` sau khi đổi.
5. **Cân nhắc**: `ExcelHelper.ts` (`any`) và `DateHelper` (class static) — đụng tới nhiều chỗ gọi, làm sau.

## 6. Cách đo lại

```sh
npx tsc --noEmit -p .                              # TS của repo (5.5.4)
npx -y -p typescript@6.0.3 tsc --noEmit -p .       # TS giống VS Code
npm run lint                                       # ESLint
```
