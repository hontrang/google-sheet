# DC

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `1291423394` — đang hiện |
| **Vùng dùng** | 2 công thức (`A1`, `A15`), mỗi cái spill `A:H` × 11 dòng |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetDC` (nhưng code projectMain không dùng) |

## Chức năng

Hiển thị danh mục top 10 cổ phiếu và phân bổ tài sản theo ngành của hai quỹ Dragon Capital — dùng làm tham chiếu "quỹ lớn đang cầm gì" (bước "Kiểm tra mã trong danh mục quỹ Dragon Capital" ở [ghi chú](ghi%20chú.md)).

| Vùng | Quỹ | Nguồn |
| --- | --- | --- |
| `A1:H12` | **DCDS** | `QUERY(IMPORTRANGE('cấu hình'!B1, "DC!A:H"), "select * limit 11")` |
| `A15:H26` | **DCDE** | `QUERY(IMPORTRANGE('cấu hình'!B1, "DC!I:P"), "select * limit 11")` |

Mỗi khối có 8 cột: Cổ phiếu, Nhóm ngành, Sàn giao dịch, Tỉ trọng NAV (%), Cập nhật lúc | Tên tài sản (ngành), Tỷ trọng, Cập nhật lúc.

Dữ liệu gốc nằm ở [data warehouse/DC](../data%20warehouse/DC.md), do custom function `LAY_THONG_TIN_DANH_MUC_DC` / `LAY_THONG_TIN_TAI_SAN_DC` sinh ra. Tab này chỉ đọc.
