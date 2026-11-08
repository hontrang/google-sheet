# Thao tác Google Drive từ Node — `driveApi.mjs`

[`src/offline/driveApi.mjs`](../src/offline/driveApi.mjs) gọi trực tiếp **Drive API v3** từ máy local: tìm, tải xuống, xuất, upload, sao chép, di chuyển, xoá và chia sẻ file trên Google Drive. Dùng được như CLI hoặc import như library. File này là anh em với [`sheetsApi.mjs`](../src/offline/sheetsApi.mjs) (đọc/ghi ô qua Sheets API) và dùng chung phần xác thực với nó.

Giống `sheetsApi.mjs`, file **không nằm trong build pipeline** — `npm run build:*` chỉ copy `src/assets/lib`, `src/utility` và folder project, nên nó không bao giờ lên GAS (xem [architecture.md §2.3](../architecture.md)). Đây là file `.mjs` chạy bằng `node` trần, không cần cài thêm package; eslint bỏ qua `**/*.mjs`.

## 1. Đăng nhập (làm một lần)

```sh
node src/offline/driveApi.mjs login              # quyền đọc + ghi
node src/offline/driveApi.mjs login --readonly   # chỉ đọc
```

Lệnh in ra một URL — mở trên trình duyệt bằng tài khoản sở hữu sheet, bấm **Cho phép**. Google redirect về `http://localhost:4571`, script bắt code và lưu refresh token vào `.sheets-token.json` ở gốc repo (chmod 600, đã có trong `.gitignore`).

| Chế độ | Scope xin cấp |
|---|---|
| mặc định | `auth/drive` + `auth/spreadsheets` |
| `--readonly` | `auth/drive.readonly` + `auth/spreadsheets.readonly` |

Điểm cần biết:

- **Một token cho cả hai thư viện.** `driveApi.mjs` và `sheetsApi.mjs` cùng đọc `.sheets-token.json`. Login của Drive luôn xin kèm scope spreadsheets, nên sau khi login ở đây, `sheetsApi.mjs` và skill `them-ma-dw` vẫn chạy bình thường. Ngược lại, `node src/offline/sheetsApi.mjs login` sẽ ghi đè token bằng bản **không có** scope drive → `driveApi.mjs` báo lỗi và cần login lại.
- OAuth client lấy từ `creds.json` trong folder `mainClasp` (biến trong `.env`), giống `sheetsApi.mjs`.
- `auth/drive` là quyền đầy đủ trên **toàn bộ** Drive của tài khoản, không chỉ hai file sheet. Chỉ cần đọc thì dùng `--readonly`.
- Token thiếu scope sẽ báo:
  ```
  LỖI: Token không có scope auth/drive. Scope hiện có: ...
  Chạy: node src/offline/driveApi.mjs login
  ```

## 2. Dùng như CLI

```sh
node src/offline/driveApi.mjs ls [<folderId>]                       # liệt kê thư mục, mặc định My Drive gốc
node src/offline/driveApi.mjs search "name contains 'bctc'"          # query Drive, xem mục 4
node src/offline/driveApi.mjs info <fileId>                          # metadata (JSON)
node src/offline/driveApi.mjs perms <fileId>                         # ai đang có quyền gì
node src/offline/driveApi.mjs download <fileId> [<file ra>]          # file thường: xlsx, pdf, ảnh…
node src/offline/driveApi.mjs export <fileId> xlsx|pdf|csv|docx [<file ra>]   # file Google
node src/offline/driveApi.mjs upload <file local> [--parent <folderId>]
node src/offline/driveApi.mjs mkdir <tên> [--parent <folderId>]
```

Chạy không tham số để in lại phần hướng dẫn ở đầu file.

CLI chỉ có lệnh **đọc** và **tạo mới**. Các thao tác sửa/xoá/chia sẻ file có sẵn chỉ gọi được qua library và bắt buộc `{ xacNhan: true }` — xem mục 3.

Ví dụ backup sheet data warehouse ra `.xlsx`:

```sh
node src/offline/driveApi.mjs export 1yttudZqXZD9URweOzEtkRDda-iUiQEh4a2uNmPlKRfw xlsx dw-backup.xlsx
```

## 3. Dùng như library

```js
import { timFile, layFileTrongThuMuc, xuatFile, taiLen, layHoacTaoThuMuc, MIME, SPREADSHEET } from './src/offline/driveApi.mjs';

// backup sheet DW vào thư mục "backup" trên Drive, dạng Google Sheet
const thuMuc = await layHoacTaoThuMuc('backup');
await xuatFile(SPREADSHEET.dw, MIME.xlsx, 'dw.xlsx');
await taiLen('dw.xlsx', { thuMucCha: thuMuc.id, ten: 'dw-2026-11-08', chuyenThanh: MIME.googleSheet });
```

Mọi hàm trả về Promise. Object file trả về có các trường `id, name, mimeType, size, modifiedTime, parents, trashed, webViewLink` (đổi được qua tuỳ chọn `truong`).

### Đọc

| Hàm | Mô tả |
|---|---|
| `timFile(q, { gioiHan = 100, truong, sapXep })` | Tìm theo [query Drive](https://developers.google.com/drive/api/guides/search-files), tự lật trang tới `gioiHan`. |
| `layFileTrongThuMuc(thuMucId = 'root')` | Con trực tiếp của thư mục, bỏ qua thùng rác. |
| `timTheoTen(ten, { thuMucCha })` | Khớp tên chính xác (đã escape `'` và `\`). |
| `layThongTin(fileId, { truong })` | Metadata, mặc định có thêm `owners`, `description`. |
| `layQuyen(fileId)` | Danh sách quyền `{ id, type, role, emailAddress, domain }`. |
| `taiXuong(fileId, dich?)` | Tải nội dung file **thường**. Có `dich` thì ghi ra đĩa, không thì trả `Buffer`. Gặp file Google thì báo lỗi, chỉ sang `xuatFile`. |
| `xuatFile(fileId, mimeType, dich?)` | Xuất file **Google** (Sheet → xlsx/pdf/csv, Doc → docx/pdf). Giới hạn 10MB của Google. Xuất `csv` chỉ lấy tab đầu tiên. |

### Tạo mới (không cần xác nhận)

| Hàm | Mô tả |
|---|---|
| `taiLen(duongDan, { ten, thuMucCha, mimeType, chuyenThanh, moTa })` | Upload multipart. MIME đoán theo đuôi file. `chuyenThanh: MIME.googleSheet` để Drive convert xlsx → Google Sheet. Luôn tạo file mới, kể cả khi trùng tên. |
| `taoThuMuc(ten, { thuMucCha })` | Tạo thư mục. |
| `layHoacTaoThuMuc(ten, { thuMucCha = 'root' })` | Có rồi thì trả về, chưa có mới tạo — dùng cho script chạy lặp lại. |
| `saoChep(fileId, { ten, thuMucCha })` | Nhân bản file (giữ nguyên công thức nếu là Google Sheet). |

### Sửa / xoá / chia sẻ (bắt buộc `{ xacNhan: true }`)

| Hàm | Mô tả |
|---|---|
| `capNhatNoiDung(fileId, duongDan, { mimeType, xacNhan })` | Ghi đè nội dung, giữ nguyên id, tên, quyền. |
| `doiTen(fileId, tenMoi, { xacNhan })` | Đổi tên. |
| `diChuyen(fileId, thuMucMoi, { xacNhan })` | Chuyển sang thư mục khác, gỡ khỏi mọi thư mục cha cũ. |
| `boVaoThungRac(fileId, { xacNhan })` | Vào thùng rác, khôi phục được trong 30 ngày. |
| `khoiPhuc(fileId)` | Lấy ra khỏi thùng rác (không cần xác nhận). |
| `xoaVinhVien(fileId, { xacNhan })` | Xoá hẳn, **không qua thùng rác**. |
| `chiaSe(fileId, { email, loai = 'user', vaiTro = 'reader', guiThongBao = false, xacNhan })` | `loai`: `user` / `group` / `domain` / `anyone`; `vaiTro`: `reader` / `commenter` / `writer`. |
| `thuHoiQuyen(fileId, quyenId, { xacNhan })` | Gỡ một quyền, `quyenId` lấy từ `layQuyen()`. |

Thiếu `xacNhan` thì hàm ném lỗi **trước khi** gửi request, ví dụ `doiTen() cần { xacNhan: true } — đây là thao tác sửa file thật trên Drive.` Đây là cùng quy ước với `ghiGiaTri` / `xoaVung` / `capNhatLo` của `sheetsApi.mjs`.

### Hằng số

- `SPREADSHEET` — ID hai sheet thật (`main`, `dw`), re-export từ `sheetsApi.mjs`. Danh sách ID và cách xác minh: [architecture.md §7](../architecture.md).
- `MIME` — `thuMuc`, `googleSheet`, `googleDoc`, `xlsx`, `docx`, `pdf`, `csv`, `json`, `txt`, `png`, `jpg`. CLI `export` nhận luôn các key này (`xlsx`, `pdf`…).

## 4. Query hay dùng

```text
name contains 'bctc' and trashed = false
'<folderId>' in parents and trashed = false
mimeType = 'application/vnd.google-apps.spreadsheet'
modifiedTime > '2026-11-01T00:00:00'
'me' in owners
```

Ghép điều kiện bằng `and` / `or`; chuỗi dùng nháy đơn. Tham chiếu đầy đủ: <https://developers.google.com/drive/api/guides/ref-search-terms>.

## 5. Ghi chú kỹ thuật

- Mọi request đều bật `supportsAllDrives=true`, còn `timFile` bật thêm `corpora=allDrives` nên thấy được cả file trong shared drive.
- Access token được cache trong RAM theo vòng đời tiến trình và kiểm tra lại scope mỗi lần dùng (`layAccessToken({ canScope })` trong `sheetsApi.mjs`).
- Lỗi API được ném ra dạng `Drive API lỗi <status>: <message>`. Lỗi `403 insufficientPermissions` khi ghi thường là do đã login bằng `--readonly`.
- Đây là bản **Node / offline**. Code chạy trong GAS muốn dùng Drive phải gọi `DriveApp` và thêm scope `https://www.googleapis.com/auth/drive` vào `appsscript.json` của project tương ứng — hiện chưa có helper nào như vậy trong `src/utility/`.
