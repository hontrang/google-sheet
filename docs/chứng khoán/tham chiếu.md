# tham chiếu

| | |
| --- | --- |
| **File** | `chứng khoán` (projectMain / Dashboard) |
| **gid** | `651281944` — đang hiện, cố định 3 dòng đầu và cột `A` |
| **Vùng dùng** | `A1:V400`, ~6.350 công thức, 11 rule định dạng có điều kiện, bộ lọc bật |
| **Hằng số trong code** | `SheetHelper.sheetName.sheetThamChieu` |

## Chức năng

Bảng quét toàn thị trường: mỗi dòng là một mã (~393 mã HOSE), mỗi cột là một chỉ số để lọc/sắp xếp — giá tham chiếu, % thay đổi, P/B, P/E, room ngoại, khối lượng, ngành, trendline, giao dịch khối ngoại, cờ "giảm N phiên liên tiếp". Dòng 1–2 là thống kê độ rộng thị trường (số mã tăng/trần/giảm/sàn/đứng giá).

## Bố cục

**Dòng 1–2 — thống kê**

| Ô | Nội dung |
| --- | --- |
| `I1` | **Ngày tham chiếu** (nhập tay, `yyyy-MM-dd`) — mốc để so giá ở cột `B`; được `cấu hình!B10` và DW `cấu hình!B1` đọc lại |
| `I2` | "cập nhật lúc" — `IMPORTRANGE` sang [data warehouse/cấu hình](../data%20warehouse/cấu%20hình.md)`!D1` (giá trị cũ, không còn được cập nhật) |
| `L2` | Thời điểm cập nhật dữ liệu ngày — `IMPORTRANGE` sang DW `cấu hình!G4` (do `layGiaKhoiLuongKhoiNgoaiMuaBanHangNgay()` ghi) |
| `O2` / `P2` / `Q2` / `R2` / `S2` | Số mã tăng (`K>0`) / trần (`K≥6,8%`) / giảm / sàn (`K≤−6,8%`) / đứng giá |
| `T2` | Ngày — DW `cấu hình!C4` |

**Dòng 3 trở xuống — một mã mỗi dòng**

| Cột | Tiêu đề | Công thức (dòng 4) |
| --- | --- | --- |
| `A` | Danh sách mã | **ghi bởi code** `getDataHose()` (xoá rồi ghi lại từ `A4`) |
| `B` | Giá tham chiếu | `XLOOKUP(A4, 'dữ liệu'!C:C, 'dữ liệu'!D:D)` |
| `C` | Chênh lệch với tham chiếu | `(giá hiện tại − B) / B` |
| `D` / `E` | P/B / P/E | `XLOOKUP` sang `dữ liệu!K` / `L` |
| `F` / `G` | Room ngoại tối đa / còn lại | `dữ liệu!M` / `N` |
| `H` | Khối lượng (TB 10 phiên) | `dữ liệu!O` |
| `I` | Ngành | `dữ liệu!H` |
| `J` | Giá ngày t-11 | `dữ liệu!AS` |
| `K` | Giá ngày t (thực chất là **% thay đổi** t so với t-1) | `(dữ liệu!BC − dữ liệu!BB) / dữ liệu!BB` |
| `L` / `M` | Trendline giá / khối lượng | `dữ liệu!DC` / `DD` (sparkline) |
| `N` | Khối ngoại ròng / khối lượng TB | `dữ liệu!DE / H` |
| `O` | Giảm N phiên liên tiếp (`ĐẠT` / `KHÔNG`) | `dữ liệu!DF` |
| `P` | Tên công ty | `dữ liệu!I` |
| `V` | Count (1 nếu mã tăng) | `IF(C4 > 0, 1, 0)` — dùng cho pivot ở bảng thông tin |

## Code liên quan

| Hàm | Trigger | Tác động |
| --- | --- | --- |
| [`getDataHose()`](../../src/projectMain/code.ts#L12) | mỗi giờ + onOpen | Xoá cột `A` từ dòng 4, ghi lại danh sách mã đọc từ `dữ liệu!A2:A` |

## Được dùng bởi

- [bảng thông tin](bảng%20thông%20tin.md): pivot ngành ở `G34` lấy nguồn `A3:V387`.
- [chi tiết mã](chi%20tiết%20mã.md)`!E16:I16`: room ngoại, KL trung bình, P/E, P/B của mã đang xem.
- [dữ liệu](dữ%20liệu.md)`!X3:X15`: % tăng/giảm hôm qua của các mã trong danh mục.
- [cấu hình](cấu%20hình.md)`!B10`: ngày tham chiếu.
- DW `cấu hình!B1` đọc ngược lại `I1` của tab này.

## Lưu ý

Các lỗi công thức đã biết — xem [sheet-audit-projectMain.md](../sheet-audit-projectMain.md): A5 (`O` truyền cả cột `A:A` làm search key), B1 (nhãn `K3`), B2 (`L:M` gần như trống), B3 (nhãn `O3` chốt cứng "9").
