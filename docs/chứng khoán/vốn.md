# vốn

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `0` (tab gốc của file) — đang hiện |
| **Vùng dùng** | `A1:H60`, 122 công thức, bộ lọc bật |
| **Hằng số trong code** | không có |

## Chức năng

Sổ nộp/rút tiền vào tài khoản chứng khoán. Mỗi dòng là một lần chuyển tiền; tab tính luỹ kế vốn đã bỏ vào, làm mẫu số cho tỷ suất lợi nhuận và % vốn mục tiêu ở [bảng thông tin](bảng%20thông%20tin.md).

## Bố cục

| Cột / ô | Tiêu đề | Nội dung |
| --- | --- | --- |
| `A` | stt | `COUNTA` luỹ kế cột `B` |
| `B` | số tiền | nhập tay (âm = rút) |
| `C` | ngày (MM/dd/yyyy) | nhập tay |
| `D` | biến động | tổng luỹ kế `SUM($B$2:B)` |
| `E2` | tiền cơ bản | vốn gốc điều chỉnh: `hằng số + SUM(B:B)` |
| `F1` | — | vốn luỹ kế mới nhất `INDEX(D:D, COUNT(D:D)+1)` |
| `H2` | phí giao dịch | `SUM('lệnh'!J:J)` |

## Được dùng bởi

- [bảng thông tin](bảng%20thông%20tin.md): `I6` (tỷ suất lợi nhuận = lãi/lỗ ÷ vốn luỹ kế), `I8` (% vốn mục tiêu = `A8 / vốn!F1`).
- [dữ liệu](dữ%20liệu.md)`!Y2`: tiền mặt = `vốn!E2` + tiền bán − tiền mua + cổ tức − `vốn!H2`.

## Lưu ý

`E2` và một số ô ở cột `B` (ví dụ `B23`) chứa hằng số điều chỉnh viết thẳng trong công thức; khi đối soát với sao kê cần xem công thức chứ không chỉ giá trị.
