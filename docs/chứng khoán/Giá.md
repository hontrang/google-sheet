# Giá

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `1497499889` — **ẩn** |
| **Vùng dùng** | 2 công thức (`A1`, `A2`), lưới 1000 × 640 |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetGia` (code projectMain không dùng tab này) |

## Chức năng

"Cửa sổ" 11 phiên gần nhất của ma trận **giá đóng cửa (VNĐ)** (dòng = ngày, cột = mã), kéo từ tab [Giá của data warehouse](../data%20warehouse/Giá.md). Ma trận gốc có hàng nghìn dòng; tab ẩn này chỉ giữ phần đuôi để các công thức `OFFSET`/`INDIRECT` ở [dữ liệu](dữ%20liệu.md) chạy trên vùng nhỏ thay vì `IMPORTRANGE` toàn bộ.

| Ô | Công thức | Ý nghĩa |
| --- | --- | --- |
| `A1` | `QUERY(IMPORTRANGE('cấu hình'!B1, "Giá!A1:OR1"), "select *")` | Dòng tiêu đề: `Ngày` + danh sách mã |
| `A2` | `QUERY(IMPORTRANGE('cấu hình'!B1, "Giá!A" & 'cấu hình'!C11 - 10 & ":OR2" & 'cấu hình'!C11), "select *")` | 11 dòng cuối của ma trận gốc; `'cấu hình'!C11` = số dòng của tab gốc |

## Được dùng bởi

[dữ liệu](dữ%20liệu.md)`!AS:BC` (giá 11 phiên của từng mã): `TRANSPOSE(OFFSET(INDIRECT("'Giá'!" & ADDRESS(COUNTA('Giá'!A:A) − …, MATCH(mã, 'Giá'!1:1, 0))), …))`.

## Lưu ý

- Chuỗi `":OR2"` là lỗi gõ (đúng ra `":OR"`), khiến range thành `…:OR2<số dòng>` (vượt xa cuối sheet). Công thức vẫn chạy chỉ vì `IMPORTRANGE` tự cắt về cuối sheet nguồn — audit mục C4.
- Dữ liệu gốc được ghi mỗi ngày bởi [`layGiaHangNgay()`](../../src/projectDw/code.ts) bên data warehouse.
