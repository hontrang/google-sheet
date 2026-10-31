# Giá

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `1178559818` — đang hiện, cố định dòng 1 và cột `A` |
| **Vùng dùng** | `A1:OQ3344` (~1,1 triệu ô giá trị), không công thức |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetGia` |

## Chức năng

Ma trận lịch sử **giá đóng cửa** (VNĐ) toàn thị trường: dòng 1 = `Ngày` + ~406 mã, mỗi dòng sau là một phiên, **cũ ở trên, mới ở dưới** (ngược với [hose](hose.md)). Dữ liệu từ khoảng năm 2013, mã niêm yết sau thì các dòng đầu để trống.

## Code liên quan

| Hàm | Trigger | Tác động |
| --- | --- | --- |
| [`layGiaKhoiLuongKhoiNgoaiMuaBanHangNgay()`](../../src/projectDw/code.ts#L44) | hằng ngày 15:50 | Gọi 4 hàm cập nhật ma trận rồi ghi `cấu hình!G4` |
| [`layGiaHangNgay()`](../../src/projectDw/code.ts#L236) | — | Nếu ô `A` dòng cuối ≠ `hose!A1`: gọi vndirect `stock_prices` cho tất cả mã ở dòng 1 (lấy từ tab **Khối Lượng**), ghi `'<ngày>` vào cột `A` dòng mới và `close × 1000` vào cột của từng mã |
| [`layGiaThamChieu()`](../../src/projectDw/code.ts#L52) | không có trigger | Đọc dòng có ngày = `cấu hình!B1` |

## Được dùng bởi

- [dữ liệu](dữ%20liệu.md)`!L:V` — giá 11 phiên gần nhất của từng mã.
- [chứng khoán/Giá](../chứng%20khoán/Giá.md) (tab ẩn) — import 11 dòng cuối; [chứng khoán/dữ liệu](../chứng%20khoán/dữ%20liệu.md)`!C1:D1` — dòng của ngày tham chiếu.
- [TRUY VAN](TRUY%20VAN.md) — một vài cột để phân tích tay.

## Lưu ý

- Thêm mã mới: phải thêm cột ở dòng 1 của **cả 4 ma trận** (Giá, Khối Lượng, KN Mua, KN Bán), vì `layGiaHangNgay()` lấy danh sách mã từ dòng 1 của tab Khối Lượng nhưng ghi theo dòng 1 của tab Giá.
- Ngày ghi dạng text (`'2026-10-12`) để so sánh chuỗi với `hose!A1`.
- Mỗi mã có giá trị ghi bằng một `setValues` riêng → ~400 lần ghi mỗi ngày, chậm nhưng nằm trong giới hạn 6 phút của GAS.
