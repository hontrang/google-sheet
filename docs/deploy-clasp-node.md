# Deploy `projectMain` — quy trình thực tế & sự cố clasp trên Node đời mới

> Bổ sung cho phần "Build & deploy" trong [CLAUDE.md](../CLAUDE.md) và mục 2.4 của [architecture.md](../architecture.md).
> File này ghi lại **cách deploy khi `npm run deploy` không chạy được**, kèm nguyên nhân gốc.

## 1. TL;DR

`npm run deploy` hiện **hỏng** trên máy này: `clasp` 2.4.2 crash ngay lúc load trên Node ≥ 24 vì `buffer.SlowBuffer` đã bị gỡ khỏi Node. Không liên quan gì tới code trong repo — `npm run build:main` vẫn chạy bình thường.

Cách lách tạm (không sửa gì trên đĩa): preload một shim rồi gọi thẳng entry point của clasp.

## 2. Nguyên nhân gốc

```
@google/clasp@2.4.2
└─ google-auth-library@7.14.1
   └─ jws@4.0.0
      └─ jwa@2.0.0
         └─ buffer-equal-constant-time@1.0.1   ← thủ phạm
```

`buffer-equal-constant-time/index.js` chạy dòng này **ở top level**, ngay khi module được load:

```js
var SlowBuffer = require('buffer').SlowBuffer;
// ...
var origSlowBufEqual = SlowBuffer.prototype.equal;   // ← TypeError
```

`SlowBuffer` là API deprecated đã bị gỡ khỏi Node đời mới. Kiểm chứng:

```console
$ node -v
v26.8.1
$ node -e "console.log(require('buffer').SlowBuffer)"
undefined
```

→ `SlowBuffer.prototype` ném `TypeError: Cannot read properties of undefined (reading 'prototype')`.

Vì nó nổ lúc **load module**, **mọi** lệnh clasp đều chết, kể cả lệnh read-only như `clasp status`:

```
/opt/homebrew/lib/node_modules/@google/clasp/node_modules/buffer-equal-constant-time/index.js:37
var origSlowBufEqual = SlowBuffer.prototype.equal;
                                  ^
TypeError: Cannot read properties of undefined (reading 'prototype')
```

## 3. Cách lách tạm (đã dùng, không sửa file nào trên đĩa)

**Bước 1** — tạo shim:

```js
// slowbuffer-shim.cjs
// Node đời mới đã gỡ buffer.SlowBuffer; clasp 2.4.2 (qua jwa/jws) vẫn đọc SlowBuffer.prototype.
// Trả lại một alias tới Buffer để module load được. Chỉ ảnh hưởng API monkeypatch deprecated
// (bufferEq.install/restore) mà clasp không dùng.
const buffer = require('buffer');
if (!buffer.SlowBuffer) buffer.SlowBuffer = buffer.Buffer;
```

**Bước 2** — build như bình thường:

```sh
npm run build:main
```

**Bước 3** — chạy clasp qua shim, **từ trong `dist/`** (clasp coi `dist/` là project root):

```sh
CLASP=$(readlink -f "$(which clasp)")   # /opt/homebrew/lib/node_modules/@google/clasp/build/src/index.js
cd dist
node --require /đường/dẫn/slowbuffer-shim.cjs "$CLASP" status   # read-only, không gọi mạng
node --require /đường/dẫn/slowbuffer-shim.cjs "$CLASP" push
```

Luôn chạy `status` trước `push`. Output đúng phải là **10 file**:

| Được push | Bị `.claspignore` loại |
| --- | --- |
| `appsscript.json` | `.clasp.json` |
| `code.ts` | `.claspignore` |
| `trigger.ts` | `.clasprc.json` |
| `utility/ChartUtil.ts` | `creds.json` |
| `utility/DateHelper.ts` | `package.json` |
| `utility/HttpHelper.ts` | `.DS_Store` |
| `utility/LogHelper.ts` | |
| `utility/SheetHelper.ts` | |
| `zAssets/Cheerio.ts` | |
| `zAssets/luxon.ts` | |

Nếu `status` in ra danh sách khác (ví dụ có `src/`, `tests/`) nghĩa là **đang đứng sai thư mục** — dừng lại, đừng push. `clasp push` gọi `script.projects.updateContent` với toàn bộ file list, file nào có trên GAS mà không có trong `dist/` sẽ **bị xoá**.

Không cần `--force`. Cờ này chỉ bỏ qua prompt xác nhận khi manifest đổi, và làm lệnh trông như thao tác phá huỷ.

## 4. Cách sửa lâu dài

| Ưu tiên | Phương án | Đánh giá |
| --- | --- | --- |
| 1 | Cài Node LTS song song (`brew install node@22`) và chạy clasp bằng Node đó | Sạch nhất. Không đụng repo, không đụng clasp. Chỉ cần trỏ đúng binary trong script `deploy` |
| 2 | Commit shim vào repo, đổi script `deploy` thành `cd dist && node --require ../src/scripts/slowbuffer-shim.cjs $(readlink -f $(which clasp)) push` | Nhanh, tự chứa trong repo. Nhược điểm: phụ thuộc vào đường dẫn cài clasp global |
| 3 | Vá thẳng file trong package clasp global | Mất mỗi lần cài lại clasp, không reproduce được trên máy khác. Chỉ dùng khi cần gấp |
| ❌ | Nâng `@google/clasp` lên v3 | **Không nên.** v3 bỏ ts2gas — không còn transpile `.ts` nữa. Toàn bộ pipeline của repo dựa vào việc clasp tự transpile `.ts` → `SERVER_JS` (xem mục 2.4 [architecture.md](../architecture.md)). Nâng lên là phải viết lại bước build |

## 5. Sau khi push — những việc push **không** tự làm

| Việc | Vì sao | Cách kích hoạt |
| --- | --- | --- |
| Vẽ lại biểu đồ `chi tiết mã` | `ChartHelper.updateChart()` chỉ chạy bên trong `layThongTinChiTietMa()` | Nhập lại mã vào `chi tiết mã!B1` (trigger `batSukienSuaThongTinO`) |
| Đăng ký trigger | `clasp push` không tạo trigger | Chạy `mainCreateTriggers` một lần trong GAS editor |
| Cập nhật Deployment ID | Push khác với deploy web app | `cấu hình!B2` phải sửa tay nếu tạo deployment mới |

## 6. Nhật ký deploy

| Ngày | Nội dung | Kết quả |
| --- | --- | --- |
| 2026-09-05 | Fix **A1** trong [sheet-audit-projectMain.md](sheet-audit-projectMain.md): `ChartUtil.ts` đổi `addRange("'dữ liệu'!AD18:AE37")` → `AE18:AE37` để series "VNINDEX" không còn vẽ nhầm cột ngày | `Pushed 10 files.` → scriptId `1LvkaTmx4VMfbQKfry_3KgU6aiQU1M7VaeMWlkp47rOyWbGel2D1GI30a`. Deploy qua shim ở mục 3 |
