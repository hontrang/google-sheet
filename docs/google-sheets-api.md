# Đọc/ghi Google Sheet từ Node — `sheetsApi.mjs`

[`src/offline/sheetsApi.mjs`](../src/offline/sheetsApi.mjs) gọi trực tiếp **Sheets API v4** lên hai sheet thật (`main` = "chứng khoán", `dw` = "data warehouse"; ID ở [architecture.md §7](../architecture.md)). Dùng được như CLI hoặc library, chạy bằng `node` trần, **không** nằm trong build GAS. Thao tác file trên Drive (tìm, backup, upload, chia sẻ) ở [`driveApi.mjs`](google-drive.md).

Khác với bản export `.xlsx`, API đọc được công thức sống, pivot table và loại lỗi của từng ô.

## 1. Đăng nhập

```sh
node src/offline/sheetsApi.mjs login        # chỉ đọc
node src/offline/sheetsApi.mjs login --rw   # đọc + ghi
```

Token lưu ở `.sheets-token.json` (gitignore, chmod 600), dùng chung với `driveApi.mjs`. Lưu ý: login ở đây **ghi đè** token và làm mất scope drive. Cần cả hai thì dùng `node src/offline/driveApi.mjs login`, lệnh này xin cả drive lẫn spreadsheets.

## 2. CLI

```sh
node src/offline/sheetsApi.mjs tabs [--id main|dw|<spreadsheetId>]
node src/offline/sheetsApi.mjs values   "'bảng thông tin'!G34:J68"   # chuỗi hiển thị
node src/offline/sheetsApi.mjs raw      "'Giá'!A1:E5" --id dw         # số thật
node src/offline/sheetsApi.mjs formulas "'tham chiếu'!V1:V10"
node src/offline/sheetsApi.mjs grid     "'bảng thông tin'!G34:J68"   # công thức + giá trị + lỗi từng ô
node src/offline/sheetsApi.mjs pivots   "bảng thông tin"
node src/offline/sheetsApi.mjs errors   "'tham chiếu'!A3:V403"
```

CLI chỉ có lệnh đọc. Không truyền `--id` thì mặc định là `main`.

## 3. Library

Mọi hàm nhận `id` (`'main'`, `'dw'` hoặc spreadsheetId đầy đủ) và trả Promise. Mọi hàm ghi/xoá/sửa cấu trúc **bắt buộc `{ xacNhan: true }`**, nếu thiếu thì ném lỗi trước khi gửi request.

### Đọc

| Hàm | Mô tả |
|---|---|
| `layDanhSachTab(id)` | `{ tenFile, tabs: [{ ten, gid, an, dong, cot }] }` |
| `layGid(tenTab, id)` | gid (sheetId) của tab, cần cho batchUpdate. Có cache. |
| `layGiaTri(range, id, { render, ngayGio })` | Mặc định là chuỗi hiển thị: `"14,500 ₫"`. `render`: `FORMATTED_VALUE` / `UNFORMATTED_VALUE` / `FORMULA`. |
| `layGiaTriTho(range, id)` | Số là `number` (`14500`), ngày vẫn là chuỗi. Dùng hàm này khi cần tính toán. |
| `layCongThuc(range, id)` | Công thức nguyên văn. |
| `layNhieuVung(ranges, { id, render = 'FORMULA', ngayGio })` | Nhiều vùng trong 1 request, trả về `{ range: values }`. |
| `layChiTietO(range, id)` | Mảng 2 chiều `{ a1, congThuc, giaTri, loaiLoi, thongBaoLoi }`. |
| `layPivot(tenTab, id)` | Pivot table trong tab. |
| `quetLoi(range, id)` | Các ô lỗi (`#REF!`, `#N/A`…). |

### Ghi giá trị

| Hàm | Mô tả |
|---|---|
| `ghiGiaTri(range, values, { id, raw, xacNhan })` | Ghi đè một vùng. Mặc định `USER_ENTERED`: chuỗi `=…` thành công thức, `"1,5"` thành số. `raw: true` để ghi nguyên văn. |
| `ghiNhieuVung(duLieu, { id, raw, xacNhan })` | Ghi nhiều vùng trong 1 request. `duLieu` là `{ range: values }` hoặc mảng `[range, values]`. |
| `themDong(range, values, { id, raw, chenDong = true, xacNhan })` | Nối dòng vào sau bảng tìm thấy trong `range` (vd. `'debug'!A:B`). Mặc định chèn dòng mới, không đè ô bên dưới. |
| `xoaVung(range, { id, xacNhan })` | Xoá nội dung, giữ định dạng. |
| `xoaNhieuVung(ranges, { id, xacNhan })` | Xoá nhiều vùng trong 1 request. |

### Cấu trúc (batchUpdate)

| Hàm | Mô tả |
|---|---|
| `capNhatLo(requests, { id, xacNhan })` | Gửi thẳng một lô [request batchUpdate](https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request). Cả lô chạy **nguyên tử**: một request lỗi thì không request nào được áp dụng. |
| `chenDongCot(tenTab, kieu, batDau, soLuong = 1, { id, keThuaTruoc, xacNhan })` | Chèn dòng/cột trống. |
| `xoaDongCot(tenTab, kieu, batDau, soLuong = 1, { id, xacNhan })` | Xoá dòng/cột, mất luôn dữ liệu trong đó. |
| `noiThemDongCot(tenTab, kieu, soLuong, { id, xacNhan })` | Nới lưới ở cuối tab. |
| `yeuCauChen(gid, …)` / `yeuCauXoa(gid, …)` / `yeuCauNoiThem(gid, …)` | Chỉ dựng object request, không gọi API. Dùng để gom nhiều thao tác vào **một** `capNhatLo()`. |

Quy ước tham số:

- `kieu` là `'cot'` hoặc `'dong'`.
- `batDau` là **index 0-based** như API: cột B = `1` (dùng `viTriCot('B')`), dòng 5 = `4`.
- Xoá nhiều cột/dòng trong cùng một tab thì xếp từ phải sang trái (dưới lên trên) để index của các phần còn lại không bị lệch.

```js
import { layDanhSachTab, capNhatLo, yeuCauXoa, viTriCot } from './src/offline/sheetsApi.mjs';

const { tabs } = await layDanhSachTab('dw');
const gid = Object.fromEntries(tabs.map((t) => [t.ten, t.gid]));
// xoá cột OS ở cả hai tab trong một lần, nguyên tử
await capNhatLo(['Giá', 'Khối Lượng'].map((t) => yeuCauXoa(gid[t], 'cot', viTriCot('OS'))), { id: 'dw', xacNhan: true });
```

### Tiện ích

- `tenCot(n)`: chuyển index thành tên cột (`407` → `'OR'`).
- `viTriCot(ten)`: chuyển tên cột thành index (`'OR'` → `407`).
- `SPREADSHEET`: map `{ main, dw }` ra ID. Ghi đè được bằng env `sheetMainId` / `sheetDwId`.

## 4. Thử lại khi lỗi

Quota của Sheets API là 60 request/phút/người dùng. Khi gặp lỗi:

- **429**: luôn thử lại, tối đa 5 lần, với backoff 2s, 4s, 8s… (cộng thêm một khoảng ngẫu nhiên nhỏ). Request đã bị từ chối nên gửi lại không gây hại.
- **5xx**: chỉ thử lại với thao tác gửi lại không gây hậu quả, gồm đọc, `ghiGiaTri`, `ghiNhieuVung`, `xoaVung`, `xoaNhieuVung`.
  - `themDong` và mọi lệnh batchUpdate **không** tự thử lại. Request có thể đã chạy xong phía server, gửi lại sẽ nối dòng hoặc xoá cột hai lần.
  - Gặp lỗi 5xx ở các hàm này thì đọc lại sheet trước khi chạy lại.

## 5. Chưa có hàm riêng

Các việc sau vẫn làm được qua `capNhatLo()` với request thô: thêm/xoá/đổi tên tab, định dạng ô, gộp ô, độ rộng cột, sắp xếp, tìm/thay thế, named range, data validation, conditional format, bảo vệ vùng. Tạo spreadsheet mới hoặc nhân bản file thì dùng `saoChep()` của [`driveApi.mjs`](google-drive.md).
