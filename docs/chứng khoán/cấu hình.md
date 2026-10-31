# cấu hình

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `396984025` — đang hiện |
| **Vùng dùng** | `A1:E23`, 16 công thức |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetCauHinh` |

## Chức năng

Bảng tham số cho cả công thức lẫn code GAS: link sang data warehouse, định dạng ngày, khoảng thời gian lấy dữ liệu, biên trục biểu đồ, từ khoá sự kiện, danh sách mã cần lấy tin, và một số credential.

## Ô tham số

| Ô | Nhãn | Giá trị / công thức | Ai dùng |
| --- | --- | --- | --- |
| `B1` | data warehouse | URL file **data warehouse** | mọi `IMPORTRANGE` trong file |
| `B2`, `D2` | Deployment ID | ID web app GAS; `D2` ghép URL `…/dev?chucnang=chiTietMa&ma=FRT` | tham khảo |
| `B3` | ABS | `abs(B16 − B17) × 2 + B4/100` trên giá `chi tiết mã` | `ChartHelper` (biên trục giá) |
| `B4` / `B5` | LOW / HIGH | `MIN` / `MAX` giá trên `chi tiết mã!B16:B35` | `ChartHelper` |
| `C3:C5` | — | Biên độ, min, max VN-INDEX từ `dữ liệu!AE18:AE37` | chưa dùng (audit D1) |
| `D4` / `D5` | — | Min / max trục VN-INDEX chốt cứng | `ChartHelper` |
| `B6` | format date | `yyyy-MM-dd` | mọi hàm code định dạng ngày |
| `B7`, `B8` | consumerID / consumerSecret | credential API SSI | không code nào đọc từ sheet (test/offline lấy từ `.env`) |
| `B9` | tài khoản | số tài khoản chứng khoán | — |
| `B10` | Ngày tham chiếu | `='tham chiếu'!I1` | `dữ liệu!D1` |
| `B11:C14` | Số hàng | `COUNTA(IMPORTRANGE(B1, "<tab>!A:A"))` cho Giá, KL, KN Mua, KN Bán | 4 tab ẩn [Giá](Giá.md), [Khối Lượng](Khối%20Lượng.md), [KN Mua](KN%20Mua.md), [KN Bán](KN%20Bán.md) |
| `B15`, `B16` | Sự kiện | `*FOMC*`, `*Fed Interest Rate*` (wildcard) | `bảng thông tin!F7:G8` |
| `B17` / `B18` | Từ ngày / Đến ngày | `TODAY() − 365` / `TODAY() + 1`, định dạng `B6` | `layGiaVaKhoiLuongTheoMaChungKhoan()`, `layHeSoBetaVaFreeFloat()` |
| `B19` | vietstock token | `__RequestVerificationToken` | ghi bởi [`vietstock()`](../../src/projectMain/code.ts#L263) |
| `B20` | derivative | kết quả thô `GetBondRelated` | ghi bởi [`derivative()`](../../src/projectMain/code.ts#L274) |
| `B21` | cột tính giá giảm liên tiếp | `='bảng thông tin'!M1` | `dữ liệu!DF`, nhãn ở `bảng thông tin!J1` |
| `B22` | Today | `YYYY/M/D` hôm nay | `bảng thông tin!C7` |
| `B23` | Tỷ lệ sheet lệnh (%) | `10` | ngưỡng giá tăng/giảm ở [lệnh](lệnh.md) `K:L` |
| `E1:E13` | danh sách tin tức | 12 mã đang theo dõi | [`layTinTucSheetBangThongTin()`](../../src/projectMain/code.ts#L37) |

## Lưu ý

- `B2`, `B7`, `B8`, `B9`, `B19` là thông tin nhạy cảm lưu dạng plain text, ai có quyền xem sheet đều đọc được — audit mục D4 đề xuất chuyển sang Script Properties.
- `B20` hiện chứa chuỗi `{RelatedBondInfos=[Ljava.lang.Object;@…}` vì `derivative()` ghi thẳng object Java vào ô; ô này không có giá trị sử dụng.
- `C11:C14` dùng offset không thống nhất (`+1` cho KN) — audit mục C6.
