# TRUY VAN

| | |
| --- | --- |
| **File** | `data warehouse` (projectDw) |
| **gid** | `1523711046` — đang hiện, cố định dòng 1 |
| **Vùng dùng** | `A1:AK2816`, ~390 công thức, 5 biểu đồ |
| **Hằng số trong code** | không có (code đọc bằng tên `'TRUY VAN'` viết thẳng) |

## Chức năng

Tab nháp để phân tích tay — rút một vài cột từ các ma trận lớn ra cạnh nhau để vẽ biểu đồ/so sánh. Không tab nào của file chứng khoán dùng tới.

| Vùng | Nội dung |
| --- | --- |
| `A1` | Một ngày đơn lẻ (`2024-08-27`) — danh sách ngày cho `duLieuTam()` |
| `E:F` | Chuỗi số + % thay đổi `(E_n − E_n−1) / E_n−1` (hiện `E` trống nên `F` toàn `#DIV/0!`) |
| `G1:K` | `QUERY` cột ngày và cột của một mã (VND) từ [Giá](Giá.md), [Khối Lượng](Khối%20Lượng.md), [KN Mua](KN%20Mua.md), [KN Bán](KN%20Bán.md) |
| `M:N` | Ngày + giá MWG từ 2013 (giá trị dán) |
| `Q:T` | Ngày (date serial) + MBB + tổng |
| `AE:AH` | `QUERY` giá VND, SSI, HCM trên dòng 2159–2777 của [Giá](Giá.md) |
| `AI:AK` | EPS của VND, SSI, HCM (giá trị dán) |

Biểu đồ: 4 line/column không tiêu đề và 1 biểu đồ cột *VND*, đều vẽ từ các vùng trên.

## Code liên quan

[`duLieuTam()`](../../src/projectDw/code.ts#L149) — không có trigger. Đọc các ngày ở `TRUY VAN!A2:A` rồi gọi 4 hàm `lay*HangNgay('Tam', ngày)` để ghi dữ liệu ngày đó vào một tab tên **`Tam`**. Dùng để backfill ngày bị thiếu; tab `Tam` hiện không tồn tại nên phải tạo tay trước khi chạy.

## Lưu ý

`F3` tham chiếu `#REF!`. Toàn bộ tab là dữ liệu thử nghiệm; có thể dọn hoặc xoá mà không ảnh hưởng luồng tự động (miễn là không chạy `duLieuTam()`).
