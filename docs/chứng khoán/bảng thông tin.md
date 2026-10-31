# bảng thông tin

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `1896535001` — đang hiện |
| **Vùng dùng** | `A1:N154`, 60 công thức, 4 biểu đồ, 1 pivot table, 15 rule định dạng có điều kiện |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetBangThongTin` |

## Chức năng

Trang tổng quan danh mục đầu tư cá nhân: NAV, lãi/lỗ, tiền mặt, tỷ trọng từng mã, sự kiện vĩ mô sắp tới, báo cáo quỹ Dragon Capital, tin tức các mã đang theo dõi và thống kê số mã tăng theo ngành. Đây là tab mở ra để xem "hôm nay danh mục thế nào".

## Bố cục

| Vùng | Nội dung | Nguồn |
| --- | --- | --- |
| `A1` | Tiêu đề | nhập tay |
| `I1` | Thời điểm cập nhật giá gần nhất | `getDataHose()` ghi qua `LogHelper.logTime` |
| `J1`, `M1` | Nhãn "Giảm N phiên liên tiếp" và số N | `J1` = `'cấu hình'!B21`; `M1` nhập tay (được `cấu hình!B21` đọc lại) |
| `J2:K10` | Danh sách mã giảm N phiên liên tiếp + tên công ty | `FILTER('dữ liệu'!A:A, 'dữ liệu'!DF:DF="ĐẠT")`, `XLOOKUP` sang `dữ liệu!I` |
| `A5:I6` | Khối chỉ số chính: Lãi, Lỗ, giá trị hiện tại, **NAV**, lãi/lỗ + cổ tức, tiền mặt, số lượng CP, tỷ suất lợi nhuận | `SUMIFS` trên [lệnh](lệnh.md), `dữ liệu!Y2`, `dữ liệu!AC3`, [cổ tức](cổ%20tức.md)`!G1`, [vốn](vốn.md) |
| `A7:A8` | Tổng vốn mục tiêu | `SUMIF` trên cột vốn mục tiêu `J23:J31` |
| `C7:C8` | Ngày đáo hạn phái sinh gần nhất | `IMPORTRANGE` → [data warehouse/dữ liệu](../data%20warehouse/dữ%20liệu.md)`!BV2` |
| `D7:E8` | 3 báo cáo mới nhất của quỹ DCDS + link | `IMPORTRANGE` → [data warehouse/DC](../data%20warehouse/DC.md)`!Q2:R4` |
| `F7:G8` | Lịch sự kiện kinh tế khớp từ khoá (mặc định `*FOMC*`, `*Fed Interest Rate*`) | `XLOOKUP` wildcard trên [Sự kiện](Sự%20kiện.md), từ khoá ở `cấu hình!B15:B16` |
| `H7:H8` | Tổng cổ tức trong năm | `cổ tức!G1` |
| `I7:I8` | % vốn mục tiêu đã đạt | `A8 / vốn!F1` |
| `A22:N31` | **Bảng danh mục**: stt, mã, số lượng, giá mua TB, giá hiện tại, giá mua thấp nhất, tăng/giảm hôm qua, giá trị, tỷ trọng, vốn mục tiêu, % vốn mục tiêu, lãi/lỗ, % lãi/lỗ, giá mua tiếp theo | `A23`, `G23`, `L23` là `OFFSET` spill từ khối danh mục ở `dữ liệu!S:AC`; `F` = `MIN(QUERY('lệnh'…))`; `N` = `F × 0,85` |
| `A34:F154` | **Tin tức**: mã, tiêu đề, link, ngày (10 tin/mã) | `layTinTucSheetBangThongTin()` cào cafef.vn cho danh sách mã ở `cấu hình!E2:E` |
| `G34:J68` | **Pivot table** theo ngành: Tổng số mã, Số mã xanh, Tỷ trọng | Nguồn `'tham chiếu'!A3:V387`, hàng = cột `I` (ngành), giá trị = `COUNTA` mã, `SUM` cột `V` (Count), công thức `sum(Count)/countA('Danh sach ma')` |

Biểu đồ: *Biểu đồ vốn đầu tư (triệu VNĐ)* (line), *Tỷ trọng cổ phiếu* (pie, nguồn `dữ liệu!T2:T20` / `Z2:Z20`), *VN-INDEX (tỷ VNĐ)* (combo, nguồn `dữ liệu!AD18:AE37`), *Vốn đầu tư (triệu đồng)* (column, nguồn `lệnh!A1:A23` / `G1:G23`).

## Code liên quan

| Hàm | Trigger | Ghi vào |
| --- | --- | --- |
| [`getDataHose()`](../../src/projectMain/code.ts#L12) | mỗi giờ + onOpen | `I1` (thời điểm cập nhật) |
| [`layTinTucSheetBangThongTin()`](../../src/projectMain/code.ts#L37) | hằng ngày 4:00 | `A35:F…` (mỗi dòng: mã, tiêu đề, `""`, link, `""`, ngày dạng text `dd/MM/yyyy`) |

## Lưu ý

- Pivot chỉ lấy nguồn tới dòng 387 của [tham chiếu](tham%20chiếu.md), trong khi danh sách mã ở đó dài tới ~dòng 396 → ~9 mã cuối bảng chữ cái không được đếm.
- Audit mục A4 trong [sheet-audit-projectMain.md](../sheet-audit-projectMain.md) kết luận khối `G35:J68` là "giá trị dán cứng" do đọc từ bản export `.xlsx`; thực tế đây là pivot table (Sheets API trả về field `pivotTable` tại `G34`). Các ô `#DIV/0!`/`#N/A` trong khối này đến từ công thức tỷ trọng tuỳ biến.
- Các lỗi công thức đã biết khác (range trôi ở `F23:F31`, `SUMIF` lệch kích thước ở `A8`, CF lệch sqref): xem mục A4, B10, B11, C1–C3 trong audit.
