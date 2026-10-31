# DC

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `1313000265` — đang hiện |
| **Vùng dùng** | `A1:S328`, 4 custom function |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetDC` |

## Chức năng

Dữ liệu quỹ **Dragon Capital** lấy từ API của dragoncapital.com.vn: danh mục top 10 cổ phiếu, phân bổ theo ngành của 2 quỹ DCDS và DCDE, và danh sách báo cáo định kỳ (NAV ngày/tuần, báo cáo tháng) của quỹ.

| Vùng | Nội dung | Nguồn |
| --- | --- | --- |
| `A1:E` | DCDS — Cổ phiếu, Nhóm ngành, Sàn, Tỉ trọng NAV (%), Cập nhật lúc | `A2 = LAY_THONG_TIN_DANH_MUC_DC('cấu hình'!B10)` |
| `F1:H` | DCDS — Tên tài sản (ngành), Tỷ trọng, Cập nhật lúc | `F2 = LAY_THONG_TIN_TAI_SAN_DC('cấu hình'!B10)` |
| `I1:M` | DCDE — như `A:E` | `I2 = LAY_THONG_TIN_DANH_MUC_DC('cấu hình'!B11)` |
| `N1:P` | DCDE — như `F:H` | `N2 = LAY_THONG_TIN_TAI_SAN_DC('cấu hình'!B11)` |
| `Q1:S328` | Tên báo cáo, Link (PDF), Cập nhật lúc — báo cáo quỹ VF1 trong năm hiện tại | **code** `layBaoCaoDC()` ghi từ `Q2` |

## Code liên quan

| Hàm | Kiểu | Ghi chú |
| --- | --- | --- |
| [`LAY_THONG_TIN_DANH_MUC_DC(url)`](../../src/projectDw/code.ts#L265) | custom function | Đọc `returnValue.top10Holding`; ngày định dạng theo `cấu hình!B6` |
| [`LAY_THONG_TIN_TAI_SAN_DC(url)`](../../src/projectDw/code.ts#L286) | custom function | Đọc `returnValue.allocationBySectors` |
| [`layBaoCaoDC()`](../../src/projectDw/code.ts#L332) | trigger **mỗi 5 phút** | POST `getDocumentContentsV2`, lấy `data[5].files` |

## Được dùng bởi

- [chứng khoán/DC](../chứng%20khoán/DC.md) — import `A:H` và `I:P` (11 dòng đầu mỗi khối).
- [chứng khoán/bảng thông tin](../chứng%20khoán/bảng%20thông%20tin.md)`!D8:E8` — 3 báo cáo mới nhất `Q2:Q4` và link `R2`.

## Lưu ý

- Custom function chỉ tính lại khi tham số đổi hoặc khi mở file, nên danh mục quỹ có thể cũ hơn ngày ở cột "Cập nhật lúc" của báo cáo.
- `layBaoCaoDC()` chọn khối báo cáo bằng chỉ số cứng `data[5]`; nếu Dragon Capital đổi thứ tự loại tài liệu, cột `Q:S` sẽ ra báo cáo khác.
- Chạy mỗi 5 phút cho một danh sách thay đổi theo ngày là thừa — tốn quota UrlFetch.
