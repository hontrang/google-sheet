# lệnh

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `689801349` — đang hiện, cố định 2 cột đầu |
| **Vùng dùng** | `A1:M53`, 366 công thức, bộ lọc bật |
| **Hằng số trong code** | không có — code GAS không đọc/ghi tab này |

## Chức năng

Nhật ký lệnh mua/bán của **năm hiện tại**. Mỗi dòng là một lệnh khớp, nhập tay ở cột `A:F`; các cột còn lại tự tính lãi/lỗ, phí và ngưỡng giá cảnh báo. Đây là nguồn gốc của gần như mọi con số danh mục ở [bảng thông tin](bảng%20thông%20tin.md) và khối danh mục trong [dữ liệu](dữ%20liệu.md).

Các năm trước được lưu ở các tab ẩn cùng cấu trúc: [2025](2025.md), [2024](2024.md), [2023](2023.md), [2022](2022.md).

## Cột

| Cột | Tiêu đề | Nhập/công thức |
| --- | --- | --- |
| `A` | ngày | nhập tay |
| `B` | mã | nhập tay |
| `C` | lệnh (`Mua` / `Bán`) | nhập tay |
| `D` | số lượng | nhập tay |
| `E` | giá mua | nhập tay |
| `F` | giá bán | nhập tay (chỉ lệnh bán) |
| `G` | tổng số tiền | `D × E` |
| `H` | tỷ lệ lãi-lỗ | Mua: `(giá hiện tại − E) / E`; Bán: `(F − E) / E` — giá hiện tại lấy từ `dữ liệu!B` |
| `I` | lãi-lỗ | `G × H` |
| `J` | phí | Mua: `E × 0,097% × D`; Bán: `E × 0,097% + F × 0,2% × D` |
| `K` | Giá tăng N% | `giá hiện tại − E × (1 + N%)`, N = `cấu hình!B23` |
| `L` | Giá giảm N% | `E × (1 − N%) − giá hiện tại` |
| `M` | Giá hiện tại | `XLOOKUP(B, 'dữ liệu'!A:A, 'dữ liệu'!B:B)` |

## Được dùng bởi

- [bảng thông tin](bảng%20thông%20tin.md): `A6` (tổng lãi), `C6` (tổng lỗ), `F6` (lãi/lỗ đã chốt), `F23:F31` (giá mua thấp nhất), biểu đồ *Vốn đầu tư*.
- [dữ liệu](dữ%20liệu.md): `T3` (`UNIQUE` các mã đã mua), `U:V` (số lượng, giá mua trung bình), `Y2` (tiền mặt).
- [vốn](vốn.md)`!H2`: tổng phí giao dịch = `SUM('lệnh'!J:J)`.

## Lưu ý

- `J5` tham chiếu `#REF!` thay vì `D5` (công thức phí lệnh bán bị hỏng ở dòng đó).
- Công thức phí lệnh bán đặt dấu ngoặc sai: `E × 0,097%` không nhân `D`, nên phí mua trong lệnh bán bị tính trên 1 cổ phiếu.
