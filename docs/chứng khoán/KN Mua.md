# KN Mua

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `275369515` — **ẩn** |
| **Vùng dùng** | 2 công thức (`A1`, `A2`), lưới 1000 × 640 |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetKhoiNgoaiMua` (code projectMain không dùng tab này) |

## Chức năng

"Cửa sổ" 21 phiên gần nhất của ma trận **khối lượng khối ngoại mua** (dòng = ngày, cột = mã), kéo từ tab [KN Mua của data warehouse](../data%20warehouse/KN%20Mua.md). Ma trận gốc có hàng nghìn dòng; tab ẩn này chỉ giữ phần đuôi để các công thức `OFFSET`/`INDIRECT` ở [dữ liệu](dữ%20liệu.md) chạy trên vùng nhỏ thay vì `IMPORTRANGE` toàn bộ.

| Ô | Công thức | Ý nghĩa |
| --- | --- | --- |
| `A1` | `QUERY(IMPORTRANGE('cấu hình'!B1, "KN Mua!A1:OR1"), "select *")` | Dòng tiêu đề: `Ngày` + danh sách mã |
| `A2` | `QUERY(IMPORTRANGE('cấu hình'!B1, "KN Mua!A" & 'cấu hình'!C13 - 20 & ":OR2" & 'cấu hình'!C13), "select *")` | 21 dòng cuối của ma trận gốc; `'cấu hình'!C13` = số dòng của tab gốc |

## Được dùng bởi

[dữ liệu](dữ%20liệu.md)`!BO:CH` (khối ngoại mua 20 phiên của từng mã): `TRANSPOSE(OFFSET(INDIRECT("'KN Mua'!" & ADDRESS(COUNTA('KN Mua'!A:A) − …, MATCH(mã, 'KN Mua'!1:1, 0))), …))`.

## Lưu ý

- Chuỗi `":OR2"` là lỗi gõ (đúng ra `":OR"`), khiến range thành `…:OR2<số dòng>` (vượt xa cuối sheet). Công thức vẫn chạy chỉ vì `IMPORTRANGE` tự cắt về cuối sheet nguồn — audit mục C4.
- Dữ liệu gốc được ghi mỗi ngày bởi [`layKhoiNgoaiMuaHangNgay()`](../../src/projectDw/code.ts) bên data warehouse.
