# debug

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `902746021` — đang hiện |
| **Vùng dùng** | `A1:B7`, không công thức |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetDebug` |

## Chức năng

Nhật ký lỗi khi chạy script: cột `A` thời gian, cột `B` nội dung, ghi bởi [`LogHelper.logDebug()`](../../src/utility/LogHelper.ts#L7). Không có dòng tiêu đề.

Nội dung hiện có: 7 dòng `Exception: Address unavailable` từ một phiên bản code cũ.

## Lưu ý

Code hiện tại không còn gọi `logDebug` ở đâu, nên tab không nhận thêm dòng mới — xem lỗi trong Execution log của GAS. Giống hệt tab [debug của chứng khoán](../chứng%20khoán/debug.md).
