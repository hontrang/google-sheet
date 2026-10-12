# Sự kiện

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `1873909790` — đang hiện, cố định dòng 1 |
| **Vùng dùng** | 1 custom function ở `A2`, spill 7 cột (~1.100 dòng) |
| **Hằng số trong code** | không có (code không ghi trực tiếp) |

## Chức năng

Lịch sự kiện kinh tế vĩ mô (chủ yếu Mỹ) với giờ đã đổi sang giờ Việt Nam. Toàn bộ dữ liệu sinh ra từ một ô:

```
A2 = LAY_SU_KIEN()
```

| Cột | Tiêu đề |
| --- | --- |
| `A` | ngày — `Thứ Ba 2026/10/13 07:15 CH` |
| `B` | quốc gia |
| `C` | tên sự kiện |
| `D` | kết quả |
| `E` | trước đó |
| `F` | nhất trí (consensus) |
| `G` | dự đoán (forecast) |

## Luồng dữ liệu

1. Workflow CI [`tradingeconomics.yml`](../../.github/workflows/tradingeconomics.yml) chạy hằng ngày, cào tradingeconomics.com/calendar và publish HTML lên GitHub Pages (`hontrang.github.io/google-sheet/`) — chi tiết ở [publish-workflow.md](../publish-workflow.md).
2. [`LAY_SU_KIEN()`](../../src/projectDw/code.ts#L305) tải trang đó, parse `table#calendar` bằng Cheerio, đổi định dạng ngày theo locale `vi-VN`.
3. Trigger `LAY_SU_KIEN` hằng ngày 1:00 tồn tại trong `dwCreateTriggers`, nhưng gọi custom function từ trigger không ghi gì vào sheet — dữ liệu trên tab thực ra được làm mới khi custom function tự tính lại (mở file / sửa ô).

## Được dùng bởi

[chứng khoán/Sự kiện](../chứng%20khoán/Sự%20kiện.md) import nguyên `A:G`, rồi [bảng thông tin](../chứng%20khoán/bảng%20thông%20tin.md)`!F8:G8` tra theo từ khoá.
