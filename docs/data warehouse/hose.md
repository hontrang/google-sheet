# hose

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `0` (tab gốc của file) — đang hiện |
| **Vùng dùng** | `A1:E6375`, ~5.500 công thức, 1 biểu đồ |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetHose` |

## Chức năng

Lịch sử chỉ số **VN-INDEX** theo ngày, **mới nhất ở dòng 1** (không có dòng tiêu đề). Đây cũng là "đồng hồ" của data warehouse: ngày ở `A1` quyết định các hàm cập nhật ma trận giá/khối lượng sẽ lấy dữ liệu ngày nào.

| Cột | Nội dung |
| --- | --- |
| `A` | Ngày (`yyyy-MM-dd`, text) |
| `B` | Điểm đóng cửa |
| `C` | % thay đổi — code ghi số cho các dòng mới; các dòng cũ là công thức `(B_n − B_n+1) / B_n+1` |
| `D` | Giá trị giao dịch (VNĐ) |
| `E` | Rải rác ~54 ô cũ ở quanh dòng 470–513 (giá trị giao dịch dạng text), là nguồn của biểu đồ line cũ trên tab |

## Code liên quan

| Hàm | Trigger | Tác động |
| --- | --- | --- |
| [`layChiSoVnIndex()`](../../src/projectDw/code.ts#L10) | **mỗi phút** | Gọi simplize `historical/quote/VNINDEX`. Nếu `A1` là hôm nay và giá trị GD đổi → ghi đè `A1:D1`; nếu là ngày mới → chèn 1 dòng ở đầu rồi ghi; nếu giá trị GD không đổi → bỏ qua |
| `layGiaHangNgay()`, `layKhoiLuongHangNgay()`, `layKhoiNgoaiMuaHangNgay()`, `layKhoiNgoaiBanHangNgay()` | 15:50 hằng ngày | Đọc `hose!A1` làm ngày cần lấy |

## Được dùng bởi

- [chứng khoán/dữ liệu](../chứng%20khoán/dữ%20liệu.md)`!AD18` — `QUERY(IMPORTRANGE(…"hose!A:D"), "select * limit 20")`: 20 phiên gần nhất cho biểu đồ VN-INDEX ở `bảng thông tin` và `chi tiết mã`.
- Các hàm cập nhật ma trận [Giá](Giá.md), [Khối Lượng](Khối%20Lượng.md), [KN Mua](KN%20Mua.md), [KN Bán](KN%20Bán.md).

## Lưu ý

- Vì so sánh theo "giá trị giao dịch có đổi không", ngày nghỉ lễ mà API vẫn trả về phiên cũ sẽ không tạo dòng mới — đúng ý đồ. Nhưng nếu API trả lỗi tạm thời rồi phục hồi với cùng ngày, dòng vẫn chỉ ghi đè, không nhân đôi.
- Không được thêm dòng tiêu đề vào tab này: code luôn đọc `A1`/`D1` như dữ liệu.
