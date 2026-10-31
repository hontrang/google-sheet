# dữ liệu

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `779304762` — đang hiện, cố định dòng 1 và cột `A` |
| **Vùng dùng** | `A1:BV407`, ~2.000 công thức, bộ lọc bật |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetDuLieu` |

## Chức năng

Bảng thông tin cơ bản theo mã của toàn thị trường (~396 mã): tên công ty, ngành, P/B, P/E, room ngoại, khối lượng trung bình, cùng chuỗi giá/khối lượng/khối ngoại các phiên gần nhất. Dashboard **chứng khoán** import `A:I` của tab này để có thông tin cơ bản mà không phải tự gọi API.

## Cột

| Cột | Tiêu đề | Nguồn |
| --- | --- | --- |
| `A` | Danh sách mã | **nhập tay** — danh sách gốc mà mọi hàm `layThongTin*` đọc (`A2:A`) |
| `B` | Tên công ty | `XLOOKUP(A, J:J, K:K)` |
| `C` | DSC | nhập tay (bản sao mã) |
| `D` | Ngành | nhập tay |
| `E` | P/B | **code** `layThongTinPB()` (vndirect `itemCode:51012`) |
| `F` | P/E | **code** `layThongTinPE()` (`itemCode:51006`) |
| `G` / `H` | Room ngoại tối đa / còn lại | **code** `layThongTinRoomNuocNgoai()` (vndirect `ownership_foreigns`) |
| `I` | Khối lượng TB 10 phiên | **code** `layThongTinKhoiLuongTrungBinh10Ngay()` (`itemCode:51016`) |
| `J` / `K` | Mã CK / Tên CTY (dữ liệu tổng) | bảng tra tên công ty, nhập tay |
| `L:V` | giá 11 phiên gần nhất | `TRANSPOSE(OFFSET(INDIRECT(… 'Giá' …)))` trên [Giá](Giá.md) |
| `W:AG` | khối lượng 11 phiên | tương tự trên [Khối Lượng](Khối%20Lượng.md) |
| `AH:BA` | nước ngoài mua 20 phiên | tương tự trên [KN Mua](KN%20Mua.md) |
| `BB:BU` | nước ngoài bán 20 phiên | tương tự trên [KN Bán](KN%20Bán.md) |
| `BV` | Phái sinh — ngày đáo hạn các HĐ tương lai chỉ số còn hiệu lực (`Thứ Năm 2026/10/15`…) | **code** `layThongTinPhaiSinh()` ghi từ `BV2` |

## Code liên quan

| Hàm | Trigger | Ghi vào |
| --- | --- | --- |
| [`layThongTinCoBan()`](../../src/projectDw/code.ts#L34) | hằng ngày 18:00 | Gọi lần lượt 4 hàm dưới, mỗi request 10 mã (`SheetHelper.kichThuocMangPhu`) |
| [`layThongTinPB()`](../../src/projectDw/code.ts#L75) | — | `E` |
| [`layThongTinPE()`](../../src/projectDw/code.ts#L94) | — | `F` |
| [`layThongTinRoomNuocNgoai()`](../../src/projectDw/code.ts#L112) | — | `G:H` |
| [`layThongTinKhoiLuongTrungBinh10Ngay()`](../../src/projectDw/code.ts#L131) | — | `I` |
| [`layThongTinPhaiSinh()`](../../src/projectDw/code.ts#L375) | hằng ngày 1:00 | `BV2:BV` |
| [`layGiaThamChieu()`](../../src/projectDw/code.ts#L52) | không có trigger | Ghi giá đóng cửa ngày `cấu hình!B1` vào cột `C` — **ghi đè cột DSC** nếu chạy tay |

## Được dùng bởi

- [chứng khoán/dữ liệu](../chứng%20khoán/dữ%20liệu.md)`!H1`, `I1`, `K1` — `QUERY(IMPORTRANGE(…"dữ liệu!A:I"))` lấy ngành, tên, P/B, P/E, room, KL TB.
- [chứng khoán/bảng thông tin](../chứng%20khoán/bảng%20thông%20tin.md)`!C8` — ngày đáo hạn phái sinh gần nhất (`BV2`).

## Lưu ý

- Dòng 138–139 có công thức `MATCH(#REF!, …)` ở `L`, `W`, `AH`, `BB` — tàn dư của việc xoá dòng/mã.
- Danh sách mã ở cột `A` phải đồng bộ thủ công với `chứng khoán/dữ liệu!A`.
