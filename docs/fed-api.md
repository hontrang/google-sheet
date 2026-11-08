# Lịch FOMC & FRED — `FomcHelper`, `FredHelper`

Hai helper trong `src/utility/` để lấy **lịch họp FOMC** và **số liệu FRED** (lãi suất mục tiêu Fed Funds…) bằng nguồn chính thức, thay cho việc cào lịch kinh tế của tradingeconomics (API của họ cần gói trả phí; tài khoản `guest` đã bị khoá — HTTP 410).

Như mọi file trong `src/utility/`, hai helper được build vào **cả** `projectMain` và `projectDw` và gọi được ở bất kỳ hàm GAS nào (một global scope chung).

| Helper | Nguồn | Cần key? |
| --- | --- | --- |
| [`FomcHelper`](../src/utility/FomcHelper.ts) | HTML `https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm` (đọc bằng Cheerio) | Không |
| [`FredHelper`](../src/utility/FredHelper.ts) | `https://api.stlouisfed.org/fred/series/observations` (JSON) | Có — miễn phí |

## `FomcHelper`

```ts
const fomc = new FomcHelper();
fomc.layLichFomc();          // IFomcMeeting[] — mọi cuộc họp trên trang (hiện 2021 → 2027), tăng dần theo ngày kết thúc
fomc.layCuocHopTiepTheo();   // cuộc họp có ngày kết thúc >= hôm nay (giờ của script)
fomc.layCuocHopGanNhat();    // cuộc họp đã kết thúc gần nhất
```

`IFomcMeeting` (`src/types/generic.ts`):

| Trường | Ví dụ | Ghi chú |
| --- | --- | --- |
| `nam` | `2026` | |
| `ngayBatDau` / `ngayKetThuc` | `2026-12-08` / `2026-12-09` | `yyyy-MM-dd`. Ngày kết thúc là ngày công bố quyết định, **14:00 giờ New York** (≈ 01:00–02:00 sáng hôm sau giờ VN) |
| `coDuBaoKinhTe` | `true` | Dấu `*` trên trang Fed — có Summary of Economic Projections (dot plot) |
| `ghiChu` | `(unscheduled)` | Rỗng nếu không có |
| `linkStatement`, `linkHopBao`, `linkBienBan` | `https://www.federalreserve.gov/newsevents/pressreleases/monetary20261028a.htm` | `undefined` khi chưa có (biên bản ra ~3 tuần sau họp) |

Trang Fed ghi tháng và ngày ở hai ô riêng (`January` + `27-28`, `Apr/May` + `30-1`, `March` + `17-18*`); `phanTichNgay()` xử lý các dạng này.

## `FredHelper`

### Lấy và lưu API key (một lần)

1. Tạo tài khoản và key miễn phí: <https://fredaccount.stlouisfed.org/apikeys> (key 32 ký tự).
2. Lưu vào **Script Properties** của project GAS, tên `FRED_API_KEY` — *không* ghi vào code hay `.env` được commit:
   - GAS editor → **Project Settings** → **Script properties** → *Add script property*; hoặc
   - chạy một lần trong editor: `FredHelper.luuApiKey('<key>')`.
3. Script Properties gắn với **từng** project GAS: muốn dùng ở cả Dashboard và Data Warehouse thì lưu ở cả hai.

### Dùng

```ts
const fred = new FredHelper();                       // đọc key từ Script Properties
fred.layLaiSuatMucTieuFed();                         // { canDuoi: 3.75, canTren: 4, ngay: '2026-11-06' }
fred.layGiaTriMoiNhat('DFEDTARU');                   // { ngay, giaTri }
fred.layQuanSat('DFEDTARU', { tuNgay: '2026-01-01', moiNhatTruoc: true, gioiHan: 30 });
```

- Ngày FRED không có số liệu (trả `"."`) bị bỏ qua.
- Lỗi API (key sai, series sai…) được **ném ra** với thông báo của FRED, ví dụ `FRED API lỗi 400: Bad Request. The value for variable api_key is not registered.` Thiếu key thì báo ngay, không gọi mạng.
- Series hay dùng: `DFEDTARU` / `DFEDTARL` (biên độ lãi suất mục tiêu), `DFF` / `EFFR` (lãi suất Fed Funds hiệu dụng), `CPIAUCSL` (CPI), `UNRATE` (thất nghiệp). Tra series khác ở <https://fred.stlouisfed.org>.

## Đang dùng ở đâu

| Hàm (projectMain) | Trigger | Làm gì |
| --- | --- | --- |
| `capNhatSuKienFed()` | hằng ngày ~7:00 — tạo bằng `taoTriggerCapNhatSuKienFed()` (cũng được gọi trong `mainCreateTriggers`) | Ghi `bảng thông tin!F7:G8`: `F7`/`F8` = `Họp FOMC sắp tới` / `08–09/12/2026 (có SEP)`; `G7`/`G8` = `Công bố lãi suất sắp tới (giờ VN)` / `02:00 10/12/2026 - hiện tại 3.75%–4.00% (FRED 07/11/2026)`. Một nguồn lỗi thì vẫn ghi phần còn lại kèm chú thích lỗi. |
| `caiDatFredApiKey(apiKey)` | chạy tay: `clasp run caiDatFredApiKey --params '["<key>"]'` | Lưu `FRED_API_KEY` vào Script Properties rồi đọc thử. Có ở cả projectMain, projectDw, projectTest. |

`FomcHelper.gioCongBoQuyetDinh(cuocHop)` đổi 14:00 giờ New York ngày cuối cuộc họp sang giờ VN (luxon tự tính giờ mùa hè/đông của Mỹ: tháng 12 → 02:00, tháng 7 → 01:00 sáng hôm sau).

Trạng thái (2026-11-08): `FRED_API_KEY` đã lưu ở cả 3 project GAS (Dashboard, Data Warehouse, Test); cả 3 đã gắn GCP project `graphite-girder-374314` nên `clasp run` chạy được.

## Test

- **QUnit trong GAS**: [`tests/gas/TestFed.ts`](../tests/gas/TestFed.ts), đã đăng ký trong `TESTS_` của `src/projectTest/code.ts` → `npm run build:test:deploy`, mở web app test. Test `layLaiSuatMucTieuFed` tự bỏ qua nếu project test chưa có `FRED_API_KEY`.
- Phần phân tích (`phanTichNgay`, `phanTichLichFomc`, `phanTichQuanSat`) tách riêng khỏi phần gọi mạng nên test được với dữ liệu cố định.

## Liên quan

- Luồng cũ (cào lịch kinh tế tradingeconomics → GitHub Pages → `LAY_SU_KIEN()` → tab Sự kiện → XLOOKUP) **đã gỡ 2026-11-08**: xoá tab Sự kiện ở cả hai sheet, xoá `LAY_SU_KIEN()` và trigger của nó, tắt workflow CI — xem [publish-workflow.md](publish-workflow.md). Mô tả chi tiết luồng cũ còn trong lịch sử git của `docs/data warehouse/Sự kiện.md`.
- [chứng khoán/bảng thông tin](chứng%20khoán/bảng%20thông%20tin.md) — ô `G8` và công thức cũ để khôi phục.
