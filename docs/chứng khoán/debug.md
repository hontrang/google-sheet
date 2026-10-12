# debug

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `534969850` — đang hiện |
| **Vùng dùng** | `A1:B22`, không công thức, bộ lọc bật |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetDebug` |

## Chức năng

Nhật ký lỗi khi chạy script: cột `A` thời gian, cột `B` nội dung. Mỗi lần gọi `LogHelper.logDebug(content)` sẽ ghi thêm một dòng vào cuối tab (đồng thời `Logger.log`).

Nội dung hiện có là lỗi mạng cũ của `getDataHose` (timeout / "Address unavailable" khi gọi `bgapidatafeed.vps.com.vn`) và lỗi `ReferenceError: OPTIONS_GET is not defined` từ một phiên bản code trước.

## Code liên quan

- [`LogHelper.logDebug()`](../../src/utility/LogHelper.ts#L7) — hàm ghi duy nhất vào tab này.

## Lưu ý

Trong code hiện tại **không còn chỗ nào gọi `logDebug`**, nên tab không nhận thêm dòng mới; lỗi runtime giờ chỉ xem được trong Execution log của GAS. Có thể xoá nội dung cũ mà không ảnh hưởng gì.
