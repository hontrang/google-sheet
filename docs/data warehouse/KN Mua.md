# KN Mua

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `1078688344` — đang hiện, cố định dòng 1 và cột `A` |
| **Vùng dùng** | `A1:OQ1993` (~790 nghìn ô giá trị), không công thức |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetKhoiNgoaiMua` |

## Chức năng

Ma trận lịch sử **khối lượng khối ngoại mua (cổ phiếu)** toàn thị trường: dòng 1 = `Ngày` + ~406 mã, mỗi dòng sau là một phiên, **cũ ở trên, mới ở dưới**. Dữ liệu từ 2018-08-30. Cùng cấu trúc với [Giá](Giá.md) (Giá có lịch sử dài hơn, từ 2013).

## Code liên quan

| Hàm | Trigger | Tác động |
| --- | --- | --- |
| [`layGiaKhoiLuongKhoiNgoaiMuaBanHangNgay()`](../../src/projectDw/code.ts#L44) | hằng ngày 15:50 | Gọi hàm bên dưới cùng 3 hàm cập nhật ma trận khác |
| [`layKhoiNgoaiMuaHangNgay()`](../../src/projectDw/code.ts#L185) | — | Nếu ô `A` dòng cuối ≠ `hose!A1`: gọi vndirect `foreigns` cho các mã ở dòng 1, ghi `'<ngày>` vào cột `A` dòng mới và trường `buyVol` vào cột của từng mã |

## Được dùng bởi

- [dữ liệu](dữ%20liệu.md)`!AH:BA` — nước ngoài mua 20 phiên của từng mã.
- [chứng khoán/KN Mua](../chứng%20khoán/KN%20Mua.md) (tab ẩn) — import các dòng cuối, phục vụ [chứng khoán/dữ liệu](../chứng%20khoán/dữ%20liệu.md)`!BO:CH`.
- [TRUY VAN](TRUY%20VAN.md) — một cột để phân tích tay.

## Lưu ý

- Thêm mã mới phải thêm cột ở dòng 1 của cả 4 ma trận (Giá, Khối Lượng, KN Mua, KN Bán).
- Ngày nào API không trả dữ liệu thì không có dòng mới, nên số dòng giữa 4 ma trận có thể lệch nhau vài phiên — các công thức đọc "N dòng cuối" vì thế có thể lệch ngày giữa giá và khối lượng.
