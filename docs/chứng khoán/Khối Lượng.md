# Khối Lượng

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `974100484` — **ẩn** |
| **Vùng dùng** | 2 công thức (`A1`, `A2`), lưới 1000 × 640 |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetKhoiLuong` (code projectMain không dùng tab này) |

## Chức năng

"Cửa sổ" 11 phiên gần nhất của ma trận **khối lượng khớp lệnh** (dòng = ngày, cột = mã), kéo từ tab [Khối Lượng của data warehouse](../data%20warehouse/Khối%20Lượng.md). Ma trận gốc có hàng nghìn dòng; tab ẩn này chỉ giữ phần đuôi để các công thức `OFFSET`/`INDIRECT` ở [dữ liệu](dữ%20liệu.md) chạy trên vùng nhỏ thay vì `IMPORTRANGE` toàn bộ.

| Ô | Công thức | Ý nghĩa |
| --- | --- | --- |
| `A1` | `QUERY(IMPORTRANGE('cấu hình'!B1, "Khối Lượng!A1:OR1"), "select *")` | Dòng tiêu đề: `Ngày` + danh sách mã |
| `A2` | `QUERY(IMPORTRANGE('cấu hình'!B1, "Khối Lượng!A" & 'cấu hình'!C12 - 10 & ":OR2" & 'cấu hình'!C12), "select *")` | 11 dòng cuối của ma trận gốc; `'cấu hình'!C12` = số dòng của tab gốc |

## Được dùng bởi

[dữ liệu](dữ%20liệu.md)`!BD:BN` (khối lượng 11 phiên của từng mã): `TRANSPOSE(OFFSET(INDIRECT("'Khối Lượng'!" & ADDRESS(COUNTA('Khối Lượng'!A:A) − …, MATCH(mã, 'Khối Lượng'!1:1, 0))), …))`.

## Lưu ý

- Chuỗi `":OR2"` là lỗi gõ (đúng ra `":OR"`), khiến range thành `…:OR2<số dòng>` (vượt xa cuối sheet). Công thức vẫn chạy chỉ vì `IMPORTRANGE` tự cắt về cuối sheet nguồn — audit mục C4.
- Dữ liệu gốc được ghi mỗi ngày bởi [`layKhoiLuongHangNgay()`](../../src/projectDw/code.ts) bên data warehouse.
