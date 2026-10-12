# cổ tức

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `716815214` — đang hiện |
| **Vùng dùng** | `A1:G29`, 27 công thức, 3 rule định dạng có điều kiện, 26 data validation |
| **Hằng số trong code** | không có |

## Chức năng

Sổ ghi cổ tức đã nhận (tiền mặt và cổ phiếu) và tổng cổ tức tiền mặt trong năm hiện tại.

## Bố cục

| Cột / ô | Tiêu đề | Nội dung |
| --- | --- | --- |
| `A` | Mã | nhập tay, có dropdown lấy từ danh sách mã ở `dữ liệu!A` |
| `B` | số tiền | thường là công thức `= 500 × C` hoặc `= C × 1000` (cổ tức/CP × số lượng) |
| `C` | số lượng | số cổ phiếu được hưởng |
| `D` | Loại cổ tức | `Tiền mặt` / `Cổ phiếu` |
| `E` | ngày (MM/dd/yyyy) | ngày nhận |
| `G1` | Tổng cổ tức trong năm | `SUMPRODUCT((YEAR(E2:E) = YEAR(TODAY())) × B2:B) × 0,95` — đã trừ 5% thuế TNCN |

## Được dùng bởi

- [bảng thông tin](bảng%20thông%20tin.md): `F6` (lãi/lỗ + cổ tức), `H8` (ô CỔ TỨC).
- [dữ liệu](dữ%20liệu.md)`!Y2`: cộng vào tiền mặt.

## Lưu ý

- `G1` cộng cả dòng cổ phiếu nếu cột `B` có giá trị — chỉ nên điền `B` cho cổ tức tiền mặt.
- 26 data validation ở cột `A` có điểm bắt đầu khác nhau do kéo range tương đối — audit mục C7.
